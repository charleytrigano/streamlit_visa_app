# =========================
# 🛂 Visas — Edition DIRECTE du fichier (app.py)
# Basé sur les préférences utilisateur (2025-10-04)
# - Edition directe d'un fichier Excel (même fichier)
# - CRUD Clients
# - Journal des paiements (add-payment) jusqu'à épuisement du solde
# - Devise USD
# - Colonne Month visible (MM uniquement), Year/MoY cachées ailleurs
# - RFE seulement si Sent/Refused/Cancelled est coché
# - Dashboard compact (KPIs + filtres)
# - Analyses avec noms détaillés (sans camembert)
# - Page Détail client (KPI, status chips, historique des paiements, ajout paiement, export CSV)
# - Widgets de date sécurisés (_date_for_widget)
# =========================

import io
import os
import uuid
from io import BytesIO
from datetime import date, datetime
from typing import List, Dict, Any

import pandas as pd
import streamlit as st

# =========================
# Utilitaires Date
# =========================

def _date_for_widget(val):
    """Convertit proprement une valeur Excel/pandas en date utilisable dans Streamlit."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return date.today()
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return date.today()
        return d.date()
    except Exception:
        return date.today()


def safe_date_input(container, label, value, key):
    return container.date_input(label, value=_date_for_widget(value), key=key, format="YYYY-MM-DD")

# =========================
# Constantes
# =========================
APP_TITLE = "🛂 Visa Manager"
CLIENTS_SHEET = "Clients"
PAYMENTS_SHEET = "Payments"
FILE_KEY = "excel_file_bytes"
FILE_NAME_KEY = "excel_file_name"
DF_CLIENTS_KEY = "df_clients"
DF_PAYMENTS_KEY = "df_payments"

CURRENCY = "USD"

# Schémas de colonnes par défaut
CLIENTS_COLUMNS = [
    "ClientID",              # str/uuid
    "FullName",              # str
    "Email",                 # str
    "Phone",                 # str
    "CreatedAt",             # date
    "Month",                 # MM (string 01..12)
    "Year",                  # YYYY (int)
    "Service",               # str (type de visa)
    "Status_Sent",           # bool
    "Status_Refused",        # bool
    "Status_Cancelled",      # bool
    "RFE",                   # bool (contrainte UI)
    "TotalAmount",           # float USD
    "PaidAmount",            # float USD
    "Balance",               # float USD (Total - Paid)
    "Notes",                 # str
]

PAYMENTS_COLUMNS = [
    "PaymentID",   # str/uuid
    "ClientID",    # fk
    "Date",        # date
    "Amount",      # float USD
    "Method",      # str
    "Reference",   # str
    "Comment",     # str
]

STATUS_CHIP_ORDER = [
    ("Status_Sent", "Sent"),
    ("Status_Refused", "Refused"),
    ("Status_Cancelled", "Cancelled"),
    ("RFE", "RFE"),
]

# =========================
# Helpers de DataFrame
# =========================

def _ensure_columns(df: pd.DataFrame, columns: List[str]) -> pd.DataFrame:
    for c in columns:
        if c not in df.columns:
            df[c] = pd.Series([None] * len(df))
    # Réordonner
    return df[columns]


def _new_client_row() -> Dict[str, Any]:
    today = date.today()
    return {
        "ClientID": str(uuid.uuid4()),
        "FullName": "",
        "Email": "",
        "Phone": "",
        "CreatedAt": today,
        "Month": f"{today.month:02d}",
        "Year": today.year,
        "Service": "",
        "Status_Sent": False,
        "Status_Refused": False,
        "Status_Cancelled": False,
        "RFE": False,
        "TotalAmount": 0.0,
        "PaidAmount": 0.0,
        "Balance": 0.0,
        "Notes": "",
    }


def _new_payment_row(client_id: str) -> Dict[str, Any]:
    return {
        "PaymentID": str(uuid.uuid4()),
        "ClientID": client_id,
        "Date": date.today(),
        "Amount": 0.0,
        "Method": "Cash",
        "Reference": "",
        "Comment": "",
    }


def _recompute_balances(df_clients: pd.DataFrame, df_payments: pd.DataFrame) -> pd.DataFrame:
    # Somme des paiements par client
    pay = df_payments.groupby("ClientID")["Amount"].sum().rename("PaidAmount").reset_index()
    df = df_clients.merge(pay, on="ClientID", how="left")
    df["PaidAmount_x"] = df["PaidAmount_x"] if "PaidAmount_x" in df.columns else None
    # Harmoniser PaidAmount
    if "PaidAmount_x" in df.columns and "PaidAmount_y" in df.columns:
        df["PaidAmount"] = df["PaidAmount_y"].fillna(0.0)
        df.drop(columns=["PaidAmount_x", "PaidAmount_y"], inplace=True)
    elif "PaidAmount_y" in df.columns:
        df["PaidAmount"] = df["PaidAmount_y"].fillna(0.0)
        df.drop(columns=[c for c in ["PaidAmount_x", "PaidAmount_y"] if c in df.columns], inplace=True)
    else:
        df["PaidAmount"] = df.get("PaidAmount", 0).fillna(0.0)

    df["TotalAmount"] = pd.to_numeric(df["TotalAmount"], errors="coerce").fillna(0.0)
    df["PaidAmount"] = pd.to_numeric(df["PaidAmount"], errors="coerce").fillna(0.0)
    df["Balance"] = (df["TotalAmount"] - df["PaidAmount"]).round(2)
    return df


# =========================
# Lecture / Ecriture Excel (openpyxl)
# =========================

def load_excel_to_session(uploaded_file: BytesIO, filename: str):
    st.session_state[FILE_NAME_KEY] = filename
    data = uploaded_file.read()
    st.session_state[FILE_KEY] = data

    with pd.ExcelFile(BytesIO(data)) as xls:
        if CLIENTS_SHEET in xls.sheet_names:
            dfc = pd.read_excel(xls, CLIENTS_SHEET)
        else:
            dfc = pd.DataFrame(columns=CLIENTS_COLUMNS)
        if PAYMENTS_SHEET in xls.sheet_names:
            dfp = pd.read_excel(xls, PAYMENTS_SHEET)
        else:
            dfp = pd.DataFrame(columns=PAYMENTS_COLUMNS)

    dfc = _ensure_columns(dfc, CLIENTS_COLUMNS)
    dfp = _ensure_columns(dfp, PAYMENTS_COLUMNS)

    # Cast simples
    dfc["CreatedAt"] = pd.to_datetime(dfc["CreatedAt"], errors="coerce").dt.date
    dfp["Date"] = pd.to_datetime(dfp["Date"], errors="coerce").dt.date

    # Recalcul soldes
    dfc = _recompute_balances(dfc, dfp)

    st.session_state[DF_CLIENTS_KEY] = dfc
    st.session_state[DF_PAYMENTS_KEY] = dfp


def save_session_to_excel_bytes() -> bytes:
    dfc = st.session_state.get(DF_CLIENTS_KEY, pd.DataFrame(columns=CLIENTS_COLUMNS))
    dfp = st.session_state.get(DF_PAYMENTS_KEY, pd.DataFrame(columns=PAYMENTS_COLUMNS))

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dfc.to_excel(writer, index=False, sheet_name=CLIENTS_SHEET)
        dfp.to_excel(writer, index=False, sheet_name=PAYMENTS_SHEET)
    return output.getvalue()


def persist_back_to_same_file():
    # Si on a un fichier téléversé initialement, on met à jour le buffer en mémoire
    data = save_session_to_excel_bytes()
    st.session_state[FILE_KEY] = data
    st.success("Fichier mis à jour en mémoire. Utilisez 'Télécharger' pour récupérer la nouvelle version.")


# =========================
# UI: Sidebar
# =========================

def sidebar_file_section():
    st.sidebar.header("Fichier Excel")
    up = st.sidebar.file_uploader("Charger le fichier Excel (xlsx)", type=["xlsx"], key="uploader")
    if up is not None:
        load_excel_to_session(up, up.name)
        st.sidebar.success(f"Chargé: {up.name}")

    if st.session_state.get(FILE_KEY):
        st.sidebar.download_button(
            "📥 Télécharger la version actuelle",
            data=st.session_state[FILE_KEY],
            file_name=st.session_state.get(FILE_NAME_KEY, "visa_manager.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if st.sidebar.button("💾 Enregistrer (écraser en mémoire)"):
            persist_back_to_same_file()


# =========================
# UI: Dashboard
# =========================

def page_dashboard():
    st.subheader("Tableau de bord")

    dfc = st.session_state.get(DF_CLIENTS_KEY, pd.DataFrame(columns=CLIENTS_COLUMNS))
    dfp = st.session_state.get(DF_PAYMENTS_KEY, pd.DataFrame(columns=PAYMENTS_COLUMNS))

    # Filtres
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        year_filter = st.selectbox("Année", ["Toutes"] + sorted([int(y) for y in set(dfc["Year"].dropna().astype(int))])) if len(dfc) else "Toutes"
    with c2:
        month_filter = st.selectbox("Mois (MM)", ["Tous"] + [f"{m:02d}" for m in range(1, 13)])
    with c3:
        service_filter = st.text_input("Service contient...", "")
    with c4:
        only_open = st.checkbox("Solde > 0", value=False)

    dff = dfc.copy()
    if year_filter != "Toutes":
        dff = dff[dff["Year"] == int(year_filter)]
    if month_filter != "Tous":
        dff = dff[dff["Month"] == month_filter]
    if service_filter:
        dff = dff[dff["Service"].fillna("").str.contains(service_filter, case=False, na=False)]
    if only_open:
        dff = dff[dff["Balance"].fillna(0) > 0]

    # KPIs
    total_clients = len(dff)
    total_amount = float(dff["TotalAmount"].fillna(0).sum())
    paid_amount = float(dff["PaidAmount"].fillna(0).sum())
    balance = float(dff["Balance"].fillna(0).sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Clients", total_clients)
    k2.metric(f"Montant Total ({CURRENCY})", f"{total_amount:,.2f}")
    k3.metric(f"Payé ({CURRENCY})", f"{paid_amount:,.2f}")
    k4.metric(f"Solde ({CURRENCY})", f"{balance:,.2f}")

    # Aperçu table (Month visible, pas Year)
    show_cols = [
        "ClientID", "FullName", "Email", "Phone",
        "CreatedAt", "Month", "Service",
        "TotalAmount", "PaidAmount", "Balance",
        "Status_Sent", "Status_Refused", "Status_Cancelled", "RFE",
    ]
    show_cols = [c for c in show_cols if c in dff.columns]

    st.dataframe(dff[show_cols].sort_values(by=["CreatedAt"], ascending=False), use_container_width=True)


# =========================
# UI: CRUD Clients
# =========================

def page_clients():
    st.subheader("Clients — CRUD")
    dfc = st.session_state.get(DF_CLIENTS_KEY, pd.DataFrame(columns=CLIENTS_COLUMNS))

    # Ajout nouveau client
    with st.expander("➕ Ajouter un client", expanded=False):
        new = _new_client_row()
        c1, c2 = st.columns(2)
        with c1:
            new["FullName"] = st.text_input("Nom complet", "")
            new["Email"] = st.text_input("Email", "")
            new["Phone"] = st.text_input("Téléphone", "")
            new["Service"] = st.text_input("Service", "")
            new["TotalAmount"] = st.number_input(f"Montant Total ({CURRENCY})", min_value=0.0, value=0.0, step=50.0)
        with c2:
            new["CreatedAt"] = safe_date_input(st, "Date de création", date.today(), key="new_created")
            # Month affiché, Year calculé
            m = st.selectbox("Mois", [f"{i:02d}" for i in range(1, 13)], index=date.today().month - 1)
            new["Month"] = m
            new["Year"] = new["CreatedAt"].year
            s1, s2, s3 = st.columns(3)
            with s1:
                new["Status_Sent"] = st.checkbox("Sent", value=False)
            with s2:
                new["Status_Refused"] = st.checkbox("Refused", value=False)
            with s3:
                new["Status_Cancelled"] = st.checkbox("Cancelled", value=False)

            # RFE seulement si au moins une des 3 est True
            rfe_enabled = new["Status_Sent"] or new["Status_Refused"] or new["Status_Cancelled"]
            new["RFE"] = st.checkbox("RFE", value=False, disabled=not rfe_enabled, help="Activable seulement si Sent/Refused/Cancelled")

        if st.button("Créer le client"):
            row = pd.DataFrame([new])
            dfc = pd.concat([dfc, row], ignore_index=True)
            dfc = _recompute_balances(dfc, st.session_state.get(DF_PAYMENTS_KEY, pd.DataFrame(columns=PAYMENTS_COLUMNS)))
            st.session_state[DF_CLIENTS_KEY] = dfc
            persist_back_to_same_file()
            st.success("Client créé.")

    st.markdown("---")

    # Edition / suppression
    if len(dfc) == 0:
        st.info("Aucun client. Ajoutez-en un.")
        return

    # Sélection client
    options = dfc.sort_values("CreatedAt", ascending=False)
    label_map = {row.ClientID: f"{row.FullName} — {row.Service} — {row.Month}/{row.Year} — {row.Balance:.2f} {CURRENCY}" for _, row in options.iterrows()}
    selected_id = st.selectbox("Sélectionner un client", list(label_map.keys()), format_func=lambda k: label_map[k])

    row_idx = dfc.index[dfc["ClientID"] == selected_id][0]
    row = dfc.loc[row_idx].to_dict()

    with st.form("edit_client_form", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            row["FullName"] = st.text_input("Nom complet", row.get("FullName", ""))
            row["Email"] = st.text_input("Email", row.get("Email", ""))
            row["Phone"] = st.text_input("Téléphone", row.get("Phone", ""))
            row["Service"] = st.text_input("Service", row.get("Service", ""))
            row["TotalAmount"] = st.number_input(f"Montant Total ({CURRENCY})", min_value=0.0, value=float(row.get("TotalAmount", 0.0)), step=50.0)
        with c2:
            row["CreatedAt"] = safe_date_input(st, "Date de création", row.get("CreatedAt"), key="edit_created")
            # Month (MM), Year synchrone
            month_current = row.get("Month") or f"{row['CreatedAt'].month:02d}"
            row["Month"] = st.selectbox("Mois", [f"{i:02d}" for i in range(1, 13)], index=int(month_current) - 1)
            row["Year"] = row["CreatedAt"].year

            s1, s2, s3, s4 = st.columns(4)
            with s1:
                row["Status_Sent"] = st.checkbox("Sent", bool(row.get("Status_Sent", False)))
            with s2:
                row["Status_Refused"] = st.checkbox("Refused", bool(row.get("Status_Refused", False)))
            with s3:
                row["Status_Cancelled"] = st.checkbox("Cancelled", bool(row.get("Status_Cancelled", False)))
            with s4:
                rfe_enabled = row["Status_Sent"] or row["Status_Refused"] or row["Status_Cancelled"]
                row["RFE"] = st.checkbox("RFE", bool(row.get("RFE", False)), disabled=not rfe_enabled)

        e1, e2, e3 = st.columns([1,1,2])
        with e1:
            save_btn = st.form_submit_button("💾 Enregistrer")
        with e2:
            delete_btn = st.form_submit_button("🗑️ Supprimer")
        with e3:
            pass

    if 'save_btn' in locals() and save_btn:
        for k, v in row.items():
            dfc.at[row_idx, k] = v
        # Recalcule des soldes
        dfc = _recompute_balances(dfc, st.session_state.get(DF_PAYMENTS_KEY, pd.DataFrame(columns=PAYMENTS_COLUMNS)))
        st.session_state[DF_CLIENTS_KEY] = dfc
        persist_back_to_same_file()
        st.success("Modifications enregistrées.")

    if 'delete_btn' in locals() and delete_btn:
        # Supprimer client + paiements associés
        dfp = st.session_state.get(DF_PAYMENTS_KEY, pd.DataFrame(columns=PAYMENTS_COLUMNS))
        dfp = dfp[dfp["ClientID"] != selected_id]
        dfc = dfc[dfc["ClientID"] != selected_id]
        st.session_state[DF_CLIENTS_KEY] = dfc.reset_index(drop=True)
        st.session_state[DF_PAYMENTS_KEY] = dfp.reset_index(drop=True)
        persist_back_to_same_file()
        st.success("Client et paiements supprimés.")


# =========================
# UI: Paiements
# =========================

def page_paiements():
    st.subheader("Paiements — Journal")
    dfc = st.session_state.get(DF_CLIENTS_KEY, pd.DataFrame(columns=CLIENTS_COLUMNS))
    dfp = st.session_state.get(DF_PAYMENTS_KEY, pd.DataFrame(columns=PAYMENTS_COLUMNS))

    if len(dfc) == 0:
        st.info("Aucun client.")
        return

    # Sélection client
    options = dfc.sort_values(["Balance", "CreatedAt"], ascending=[False, False])
    label_map = {row.ClientID: f"{row.FullName} — Solde {row.Balance:.2f} {CURRENCY}" for _, row in options.iterrows()}
    selected_id = st.selectbox("Client", list(label_map.keys()), format_func=lambda k: label_map[k])

    client = dfc[dfc["ClientID"] == selected_id].iloc[0]

    st.write(f"**Total**: {client.TotalAmount:.2f} {CURRENCY} | **Payé**: {client.PaidAmount:.2f} {CURRENCY} | **Solde**: {client.Balance:.2f} {CURRENCY}")

    with st.form("add_payment_form"):
        newp = _new_payment_row(selected_id)
        newp["Date"] = safe_date_input(st, "Date du paiement", date.today(), key="pay_date")
        max_amount = max(0.0, float(client.Balance))
        newp["Amount"] = st.number_input(f"Montant ({CURRENCY})", min_value=0.0, max_value=float(max_amount), value=float(max_amount) if max_amount > 0 else 0.0, step=10.0, help="Le montant est plafonné au solde restant.")
        newp["Method"] = st.selectbox("Méthode", ["Cash", "Card", "Wire", "Other"]) 
        newp["Reference"] = st.text_input("Référence", "")
        newp["Comment"] = st.text_input("Commentaire", "")
        add_btn = st.form_submit_button("Ajouter le paiement")

    if add_btn:
        if newp["Amount"] <= 0.0:
            st.warning("Montant invalide.")
        else:
            dfp = pd.concat([dfp, pd.DataFrame([newp])], ignore_index=True)
            st.session_state[DF_PAYMENTS_KEY] = dfp
            # Recalcule des soldes
            dfc = _recompute_balances(dfc, dfp)
            st.session_state[DF_CLIENTS_KEY] = dfc
            persist_back_to_same_file()
            st.success("Paiement ajouté.")

    # Historique du client
    hist = dfp[dfp["ClientID"] == selected_id].sort_values("Date", ascending=False)
    st.markdown("### Historique des paiements")
    st.dataframe(hist, use_container_width=True)


# =========================
# UI: Analyses
# =========================

def page_analyses():
    st.subheader("Analyses (sans camembert)")
    dfc = st.session_state.get(DF_CLIENTS_KEY, pd.DataFrame(columns=CLIENTS_COLUMNS))
    dfp = st.session_state.get(DF_PAYMENTS_KEY, pd.DataFrame(columns=PAYMENTS_COLUMNS))

    if len(dfc) == 0:
        st.info("Aucune donnée.")
        return

    c1, c2 = st.columns(2)
    with c1:
        year = st.selectbox("Année", sorted(dfc["Year"].dropna().astype(int).unique()), index=0)
    with c2:
        month = st.selectbox("Mois (MM)", [f"{i:02d}" for i in range(1, 13)])

    dff = dfc[(dfc["Year"] == int(year)) & (dfc["Month"] == month)].copy()

    # KPIs sur période
    k1, k2, k3 = st.columns(3)
    k1.metric("Clients", len(dff))
    k2.metric(f"Total ({CURRENCY})", f"{dff['TotalAmount'].fillna(0).sum():,.2f}")
    k3.metric(f"Solde ({CURRENCY})", f"{dff['Balance'].fillna(0).sum():,.2f}")

    st.markdown("### Détail par client")
    cols = ["FullName", "Service", "CreatedAt", "Month", "TotalAmount", "PaidAmount", "Balance", "RFE"]
    st.dataframe(dff[cols].sort_values("CreatedAt", ascending=False), use_container_width=True)


# =========================
# UI: Détail client
# =========================

def page_detail_client():
    st.subheader("Fiche Client")
    dfc = st.session_state.get(DF_CLIENTS_KEY, pd.DataFrame(columns=CLIENTS_COLUMNS))
    dfp = st.session_state.get(DF_PAYMENTS_KEY, pd.DataFrame(columns=PAYMENTS_COLUMNS))

    if len(dfc) == 0:
        st.info("Aucun client.")
        return

    options = dfc.sort_values("CreatedAt", ascending=False)
    label_map = {row.ClientID: f"{row.FullName} — {row.Service} — Solde {row.Balance:.2f} {CURRENCY}" for _, row in options.iterrows()}
    selected_id = st.selectbox("Client", list(label_map.keys()), format_func=lambda k: label_map[k])

    cli = dfc[dfc["ClientID"] == selected_id].iloc[0]

    # Chips statut
    st.markdown("### Statuts")
    chips = []
    for col, label in STATUS_CHIP_ORDER:
        val = bool(cli.get(col, False))
        style = "✅" if val else "⬜"
        chips.append(f"{style} {label}")
    st.write(" | ".join(chips))

    # KPIs
    k1, k2, k3 = st.columns(3)
    k1.metric("Total", f"{cli.TotalAmount:.2f} {CURRENCY}")
    k2.metric("Payé", f"{cli.PaidAmount:.2f} {CURRENCY}")
    k3.metric("Solde", f"{cli.Balance:.2f} {CURRENCY}")

    # Historique paiements
    st.markdown("### Paiements")
    hist = dfp[dfp["ClientID"] == selected_id].sort_values("Date", ascending=False)
    st.dataframe(hist, use_container_width=True)

    # Ajouter un paiement rapide
    st.markdown("### Ajouter un paiement")
    with st.form("quick_pay_form"):
        amount = st.number_input(f"Montant ({CURRENCY})", min_value=0.0, max_value=float(max(0.0, cli.Balance)), value=float(max(0.0, cli.Balance)), step=10.0)
        method = st.selectbox("Méthode", ["Cash", "Card", "Wire", "Other"], key="quick_method")
        pdate = safe_date_input(st, "Date", date.today(), key="quick_date")
        comment = st.text_input("Commentaire", "")
        submit = st.form_submit_button("Ajouter")

    if submit:
        if amount > 0:
            newp = {
                "PaymentID": str(uuid.uuid4()),
                "ClientID": selected_id,
                "Date": pdate,
                "Amount": float(amount),
                "Method": method,
                "Reference": "",
                "Comment": comment,
            }
            dfp = pd.concat([dfp, pd.DataFrame([newp])], ignore_index=True)
            st.session_state[DF_PAYMENTS_KEY] = dfp
            # Recompute balances
            dfc = _recompute_balances(dfc, dfp)
            st.session_state[DF_CLIENTS_KEY] = dfc
            persist_back_to_same_file()
            st.success("Paiement ajouté.")
        else:
            st.warning("Montant invalide.")

    # Export CSV (fiche client + paiements)
    st.markdown("### Export")
    export_df = pd.merge(
        dfc[dfc["ClientID"] == selected_id],
        dfp[dfp["ClientID"] == selected_id],
        on="ClientID",
        how="left",
        suffixes=("_Client", "_Payment"),
    )
    csv_bytes = export_df.to_csv(index=False).encode("utf-8")
    st.download_button("Exporter CSV (fiche + paiements)", data=csv_bytes, file_name=f"client_{selected_id}.csv", mime="text/csv")


# =========================
# Main App
# =========================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    sidebar_file_section()

    tabs = st.tabs(["🏠 Dashboard", "👤 Clients", "💵 Paiements", "📊 Analyses", "🗂️ Détail client"]) 
    with tabs[0]:
        page_dashboard()
    with tabs[1]:
        page_clients()
    with tabs[2]:
        page_paiements()
    with tabs[3]:
        page_analyses()
    with tabs[4]:
        page_detail_client()


if __name__ == "__main__":
    main()
