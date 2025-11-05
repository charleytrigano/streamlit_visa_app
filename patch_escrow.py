# -*- coding: utf-8 -*-
"""
Patch Escrow pour app.py -> app_escrow.py (modification non destructive)
- Ne change PAS la logique des données (Escrow reste dans le modèle).
- Masque Escrow dans Dashboard, Analyses, Ajouter, Gestion.
- Crée un onglet "🧾 Escrow" listant: Dossier N, Nom, Date, Acompte 1, Date Acompte 1, Dossiers envoyé,
  avec export Excel.
Usage:
    python3 patch_escrow.py [app.py]  # par défaut: app.py dans le répertoire courant
"""

import sys, re
from pathlib import Path

SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("app.py")
DST = SRC.with_name("app_escrow.py")

if not SRC.exists():
    print(f"[ERREUR] Fichier source introuvable: {SRC}")
    sys.exit(1)

text = SRC.read_text(encoding="utf-8")

# 1) Injecter helpers après st.title(APP_TITLE)
helper_code = r'''
# --- helper to hide Escrow in selected UI tables (we keep the column in the data model) ---
def hide_escrow(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if isinstance(df, pd.DataFrame):
            return df.drop(columns=["Escrow"], errors="ignore")
    except Exception:
        pass
    return df

def st_display_df(df, **kwargs):
    # Use this helper to show dataframes with Escrow column removed (keeps internal data intact)
    try:
        df2 = hide_escrow(df)
    except Exception:
        df2 = df
    return st.dataframe(df2, **kwargs)
# --- end helpers ---
'''

m = re.search(r"(st\.title\s*\(\s*APP_TITLE\s*\)\s*)", text)
if m:
    idx = m.end()
    text = text[:idx] + "\n" + helper_code + text[idx:]
else:
    # fallback: tenter après set_page_config + title groupé
    m2 = re.search(r"(st\.set_page_config\(.*?\)\s*\n\s*st\.title\s*\(\s*APP_TITLE\s*\)\s*)", text, flags=re.DOTALL)
    if m2:
        idx = m2.end()
        text = text[:idx] + "\n" + helper_code + text[idx:]
    else:
        # dernier recours: préfixer
        text = helper_code + "\n" + text

# 2) Ajouter "🧾 Escrow" au tableau des tabs
def append_escrow_to_tabs(s: str) -> str:
    pat = re.compile(r"(tabs\s*=\s*st\.tabs\s*\(\s*\[)(.*?)(\]\s*\))", flags=re.DOTALL)
    m = pat.search(s)
    if not m:
        return s
    start, inner, end = m.group(1), m.group(2), m.group(3)
    if "Escrow" in inner or "🧾" in inner:
        return s
    new_inner = inner.rstrip() + ', "🧾 Escrow"'
    return s[:m.start()] + start + new_inner + end + s[m.end():]

text = append_escrow_to_tabs(text)

# 3) Définir une variable escrow_tab = tabs[-1] juste après la définition des tabs
text = re.sub(
    r"(tabs\s*=\s*st\.tabs\s*\(\s*\[.*?\]\s*\)\s*)",
    r"\1\nescrow_tab = tabs[-1]\n",
    text,
    flags=re.DOTALL
)

# 4) Remplacer st.dataframe/table par st_display_df dans les blocs tabs[1]..tabs[4]
def replace_in_tab_block(s: str, tab_index: int) -> str:
    start_pat = re.compile(rf"(with\s+tabs\[{tab_index}\]\s*:\s*)", flags=re.MULTILINE)
    m = start_pat.search(s)
    if not m:
        return s
    start_idx = m.end()
    next_tab = re.search(r"with\s+tabs\[\d+\]\s*:", s[start_idx:])
    end_idx = start_idx + next_tab.start() if next_tab else len(s)
    block = s[start_idx:end_idx]
    block2 = block.replace("st.dataframe(", "st_display_df(")
    block2 = block2.replace("st.table(", "st_display_df(")
    return s[:start_idx] + block2 + s[end_idx:]

for i in [1, 2, 3, 4]:
    text = replace_in_tab_block(text, i)

# 5) Supprimer le checkbox "Escrow" dans l'onglet Ajouter
text = re.sub(
    r'^\s*add_escrow\s*=\s*st\.checkbox\([^\)]*["\']Escrow["\'][^\)]*\)\s*\n',
    '',
    text,
    flags=re.MULTILINE
)

# 6) Supprimer l’affectation new_row["Escrow"] = ...
text = re.sub(
    r'^\s*new_row\["Escrow"\]\s*=\s*.*\n',
    '',
    text,
    flags=re.MULTILINE
)

# 7) Ajouter le bloc entier de l’onglet Escrow à la fin du fichier
escrow_tab_block = r'''
# -------------------------
# 🧾 Escrow (onglet dédié)
# -------------------------
with escrow_tab:
    st.header("🧾 Dossiers en Escrow")
    df_live_esc = recalc_payments_and_solde(_get_df_live_safe()) if 'recalc_payments_and_solde' in globals() else _get_df_live_safe()
    if df_live_esc is None or (hasattr(df_live_esc,'empty') and df_live_esc.empty):
        st.info("Aucune donnée en mémoire.")
    else:
        try:
            import pandas as pd
            from io import BytesIO
            # Filtrer Escrow == 1
            try:
                df_esc = df_live_esc[df_live_esc.get("Escrow", 0) == 1].copy()
            except Exception:
                # fallback si .get indispo
                df_esc = df_live_esc[df_live_esc["Escrow"] == 1].copy() if "Escrow" in df_live_esc.columns else df_live_esc.iloc[0:0].copy()

            display_cols = ["Dossier N","Nom","Date","Acompte 1","Date Acompte 1","Dossiers envoyé"]
            existing = [c for c in display_cols if c in df_esc.columns]
            if not existing:
                st.info("Aucun champ Escrow pertinent trouvé dans les données.")
            else:
                df_show = df_esc[existing].reset_index(drop=True)

                # Formatting dates & money
                for c in ["Date","Date Acompte 1"]:
                    if c in df_show.columns:
                        try:
                            df_show[c] = pd.to_datetime(df_show[c], errors="coerce")
                        except Exception:
                            pass
                if "Acompte 1" in df_show.columns:
                    try:
                        def _to_num(x):
                            try:
                                import re
                                s = str(x).strip().replace('\u202f','').replace('\xa0','')
                                s = re.sub(r"[^\d,.\-]", "", s)
                                if s.count(",")==1 and s.count(".")==0:
                                    # 1,23 -> 1.23
                                    if len(s.split(",")[-1]) in (2,3):
                                        s = s.replace(",", ".")
                                s = s.replace(",", "")
                                return float(s)
                            except Exception:
                                try:
                                    return float(x)
                                except Exception:
                                    return 0.0
                        def _fmt_money(v):
                            try:
                                return "${:,.2f}".format(float(v))
                            except Exception:
                                return "$0.00"
                        df_show["Acompte 1"] = df_show["Acompte 1"].apply(lambda v: _fmt_money(_to_num(v)))
                    except Exception:
                        pass

                st.dataframe(df_show, use_container_width=True, height=480)

                # Export Excel
                try:
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        df_show.to_excel(writer, index=False, sheet_name="Escrow")
                    st.download_button(
                        "Exporter Escrow en Excel",
                        data=buf.getvalue(),
                        file_name="escrow_export.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Erreur export Excel: {e}")
        except Exception as e:
            st.error(f"Erreur affichage Escrow: {e}")
'''

if "with escrow_tab:" not in text:
    text = text.rstrip() + "\n\n" + escrow_tab_block + "\n"

# 8) Écrire le fichier final
DST.write_text(text, encoding="utf-8")
print(f"[OK] Fichier généré: {DST}")
