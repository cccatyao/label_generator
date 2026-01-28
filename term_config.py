#!/usr/bin/env python3
"""
Term Dictionary Configuration for Label 19
Contains Part and Material term mappings with French translations.
"""

# Part titles - English (normalized uppercase) -> French translation
PARTS_DICT = {
    "BODY": "Le matériau principal du corps",
    "SEAT CUSHION": "Un coussin de siège",
    "BACK CUSHION": "Un coussin de dossier",
    "PILLOW BOLSTER": "Coussin pour accoudoir",
    "ARM CUSHION": "Accoudoir",
    "SUPPORT BAR": "Barre de soutien",
    "SEAT BASE": "Base de siège",
}

# Material terms - English (normalized uppercase) -> French translation
MATERIALS_DICT = {
    "POLYURETHANE FOAM PAD": "Tampon en mousse de polyuréthane",
    "POLYESTER FIBER BATTING": "La ouate de fibre de polyester",
    "COIL POCKET SPRING": "Ensemble de ressorts enroulés",
    "POLYESTER FIBER": "Fibre de polyester",
    "PVC BOARD": "Planche de PVC",
    "WATERFOWL FEATHERS": "Les plumes de canard",
    "EXPANDED POLYTHYLENE FOAM": "Mousse de polyéthylène expansé",
    "FELT":"Feutre",
    "IRON PAD":"Plaque de fixation",
}


def normalize_text(text: str) -> str:
    """Normalize text for matching: uppercase, single spaces, stripped."""
    import re
    text = text.upper().strip()
    text = re.sub(r'\s+', ' ', text)
    return text


def find_part_match(text: str) -> tuple:
    """
    Check if text contains any part keyword.
    Returns (matched_key, french_translation) or (None, None) if not found.
    """
    normalized = normalize_text(text)
    # Sort by length descending to match longer terms first (e.g., "SEAT CUSHION" before "CUSHION")
    for key in sorted(PARTS_DICT.keys(), key=len, reverse=True):
        if key in normalized:
            return key, PARTS_DICT[key]
    return None, None


def find_material_match(text: str) -> tuple:
    """
    Check if text contains any material keyword.
    Returns (matched_key, french_translation) or (None, None) if not found.
    """
    normalized = normalize_text(text)
    # Sort by length descending to match longer terms first
    for key in sorted(MATERIALS_DICT.keys(), key=len, reverse=True):
        if key in normalized:
            return key, MATERIALS_DICT[key]
    return None, None
