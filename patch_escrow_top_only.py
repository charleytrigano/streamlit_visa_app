# -*- coding: utf-8 -*-
"""
Patch : garder Escrow uniquement dans la barre d'onglets principale (en haut)
et supprimer toute ligne d'onglets séparée affichée en bas.
Usage :
    python3 patch_escrow_top_only.py [app.py]
→ crée un fichier app_escrow_top.py propre.
"""

import sys, re
from pathlib import Path

SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("app.py")
DST = SRC.with_name("app_escrow_top.py")

if not SRC.exists():
    print(f"[ERREUR] Fichier source introuvable : {SRC}")
    sys.exit(1)

text = SRC.read_text(encoding="utf-8")

# 1️⃣ Ajouter Escrow dans les onglets principaux si manquant
tabs_pattern = re.compile(r'(st\.tabs\s*\(\s*\[)([^\]]*)(\]\s*\))', re.DOTALL)
def add_escrow_to_tabs(match):
    inside = match.group(2)
    if "Escrow" not in inside:
        if inside.strip().endswith(","):
            new_inside = inside + ' "🧾 Escrow"'
        else:
            new_inside = inside.rstrip() + ', "🧾 Escrow"'
        return f"{match.group(1)}{new_inside}{match.group(3)}"
    return match.group(0)

text = re.sub(tabs_pattern, add_escrow_to_tabs, text)

# 2️⃣ Supprimer toute création secondaire de tab Escrow séparée
# Exemples ciblés : escrow_tab = st.tabs(["Escrow"]) ou st.tabs(["🧾 Escrow"])
text = re.sub(
    r'^\s*\w*\s*=\s*st\.tabs\s*\(\s*\[\s*["\']🧾?\s*Escrow["\'].*?\]\s*\)\s*.*?$',
    '',
    text,
    flags=re.MULTILINE,
)

# 3️⃣ Supprimer les blocs "with st.tabs(['Escrow']):"
text = re.sub(
    r'^\s*with\s+st\.tabs\s*\(\s*\[\s*["\']🧾?\s*Escrow["\'].*?\]\s*\)\s*:\s*.*?(?=^\S|\Z)',
    '',
    text,
    flags=re.MULTILINE | re.DOTALL,
)

# 4️⃣ Enregistrer la nouvelle version
DST.write_text(text, encoding="utf-8")
print(f"[OK] Escrow ajouté en haut, onglet du bas supprimé → {DST}")
