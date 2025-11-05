# -*- coding: utf-8 -*-
"""
Patch : retirer l’onglet Escrow du menu et masquer son bloc de code.
Usage :
    python3 patch_remove_escrow_tab.py [app.py]
→ crée un fichier app_noescrow.py sans onglet Escrow visible.
"""

import sys, re
from pathlib import Path

SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("app.py")
DST = SRC.with_name("app_noescrow.py")

if not SRC.exists():
    print(f"[ERREUR] Fichier source introuvable : {SRC}")
    sys.exit(1)

text = SRC.read_text(encoding="utf-8")

# 1️⃣ Supprimer "Escrow" dans la liste des onglets st.tabs([...])
pattern_tabs = re.compile(r'(st\.tabs\s*\(\s*\[)([^\]]*)(\]\s*\))', re.DOTALL)
def remove_escrow_from_tabs(match):
    inside = match.group(2)
    cleaned = re.sub(r'["\']\s*🧾?\s*Escrow\s*["\']\s*,?\s*', '', inside)
    return f"{match.group(1)}{cleaned}{match.group(3)}"

text = re.sub(pattern_tabs, remove_escrow_from_tabs, text)

# 2️⃣ Supprimer la définition "escrow_tab = ..." si elle existe
text = re.sub(r'^\s*escrow_tab\s*=\s*tabs\[-1\]\s*\n', '', text, flags=re.MULTILINE)

# 3️⃣ Commenter tout bloc "with escrow_tab:" (section Escrow)
pattern_escrow_block = re.compile(r'^\s*with\s+escrow_tab\s*:.*?(?=^\S|\Z)', re.DOTALL | re.MULTILINE)
text = re.sub(pattern_escrow_block,
              lambda m: "\n".join("# " + line for line in m.group(0).splitlines()),
              text)

# 4️⃣ Écriture du fichier modifié
DST.write_text(text, encoding="utf-8")
print(f"[OK] Onglet Escrow supprimé. Nouveau fichier généré : {DST}")
