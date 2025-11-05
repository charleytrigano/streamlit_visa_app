# -*- coding: utf-8 -*-
"""
Patch : masquer la détection et l'import ComptaCli dans l'onglet Fichiers.
Usage :
    python3 patch_comptacli_hide.py [app.py]
→ crée un fichier app_nocomptacli.py sans afficher le bloc ComptaCli détectée.
"""

import sys, re
from pathlib import Path

SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("app.py")
DST = SRC.with_name("app_nocomptacli.py")

if not SRC.exists():
    print(f"[ERREUR] Fichier source introuvable : {SRC}")
    sys.exit(1)

text = SRC.read_text(encoding="utf-8")

# Recherche et mise en commentaire du bloc "ComptaCli détectée"
pattern = re.compile(
    r"(?ms)^\s*if\s+.*comptacli.*detected.*?:.*?(?:st\.[a-z_]+\([^)]*\).*){1,20}"
)
patched = re.sub(pattern, lambda m: "\n".join("# " + l for l in m.group(0).splitlines()), text)

# Si rien n'a été trouvé, on affiche une info
if patched == text:
    print("[INFO] Aucun bloc ComptaCli détecté à masquer.")
else:
    print("[OK] Bloc ComptaCli détecté et masqué.")

DST.write_text(patched, encoding="utf-8")
print(f"[OK] Nouveau fichier généré : {DST}")
