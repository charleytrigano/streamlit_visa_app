# ... le reste inchangé ...

def _norm_key(s: str) -> str:
    s = "".join(c for c in unicodedata.normalize("NFKD", s or "") if not unicodedata.combining(c))
    return s.lower().strip()

def validate_rfe_row(row: pd.Series) -> Tuple[bool, str]:
    """
    Valide la cohérence des statuts d'un dossier (RFE, Envoyé, Refusé, Approuvé, Annulé).
    Supporte les variantes d'accents/espaces.
    """
    # normalise les clés une seule fois
    keys = {_norm_key(k): k for k in row.index}

    def has_any(candidates):
        for cand in candidates:
            k = keys.get(_norm_key(cand))
            if k and bool(row.get(k)):
                return True
        return False

    rfe = has_any(["RFE"])
    sent = has_any(["Dossier envoyé", "Dossier envoye"])
    refused = has_any(["Dossier refusé", "Dossier refuse"])
    approved = has_any(["Dossier approuvé", "Dossier approuve"])
    canceled = has_any(["DossierAnnule", "Dossier Annule", "Dossier annulé"])

    if rfe and not (sent or refused or approved):
        return False, "RFE doit être combinée avec Envoyé / Refusé / Approuvé"
    if approved and refused:
        return False, "Un dossier ne peut pas être à la fois Approuvé et Refusé"
    if canceled and (sent or refused or approved):
        return False, "Un dossier annulé ne peut pas être marqué Envoyé/Refusé/Approuvé"
    return True, ""

