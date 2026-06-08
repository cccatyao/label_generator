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
    "HEADBOARD EXTENSION": "Prolongement de tête de lit",
    "TOP FABRIC": "Tissu du dessus",
    "BORDER FABRIC": "Tissu de bordure",
}

# Origin country mapping
ORIGIN_COUNTRY_MAP = {
    "CN": "CHINA",
    "VN": "VIETNAM",
    "KHM": "CAMBODIA",
    "US": "UNITED STATES",
}

# Material terms - English (normalized uppercase) -> French translation
# All material terms should also be updated to LABLE4_TITLE_MAP
MATERIALS_DICT = {
    "VISCOELASTIC POLYURETHANE FOAM": "Mousse de polyuréthane viscoélastique",
    "POLYURETHANE FOAM PAD": "Tampon en mousse de polyuréthane",
    "POLYURETHANE FOAM": "Mousse de polyuréthane",
    "POLYESTER FIBER BATTING": "La ouate de fibre de polyester",
    "COIL POCKET SPRING": "Ensemble de ressorts enroulés",
    "INNERSPRING UNIT": "Unité de ressorts",
    "POLYESTER FIBER": "Fibre de polyester",
    "PVC BOARD": "Planche de PVC",
    "WATERFOWL FEATHERS": "Les plumes de canard",
    "EXPANDED POLYTHYLENE FOAM": "Mousse de polyéthylène expansé",
    "LYOCELL": "Lyocell",
    "COTTON CANVAS": "Toile de coton",
    "FELT": "Feutre",
    "IRON PAD": "Plaque de fixation",
}

# Label 4 title mapping:
# English (canonical display) -> French + Type
# Supported types: "Material", "Part Title", "Sub-part Title"
LABEL4_TITLE_MAP = {
    # Keep Label 4 material coverage in sync with MATERIALS_DICT above.
    # When adding a MATERIALS_DICT term, add a matching "Material" entry here too.
    "Viscoelastic Polyurethane Foam": {"french": "Mousse de polyuréthane viscoélastique", "type": "Material"},
    "Polyurethane Foam Pad": {"french": "Tampon en mousse de polyuréthane", "type": "Material"},
    "Polyurethane Foam": {"french": "Mousse de polyuréthane", "type": "Material"},
    "Polyester Fiber Batting": {"french": "La ouate de fibre de polyester", "type": "Material"},
    "Coil Pocket Spring": {"french": "Ensemble de ressorts enroulés", "type": "Material"},
    "Innerspring Unit": {"french": "Unité de ressorts", "type": "Material"},
    "Polyester Fiber": {"french": "Fibre de polyester", "type": "Material"},
    "PVC Board": {"french": "Planche de PVC", "type": "Material"},
    "Waterfowl Feathers": {"french": "Les plumes de canard", "type": "Material"},
    "Expanded Polythylene Foam": {"french": "Mousse de polyéthylène expansé", "type": "Material"},
    "Lyocell": {"french": "Lyocell", "type": "Material"},
    "Cotton Canvas": {"french": "Toile de coton", "type": "Material"},
    "Felt": {"french": "Feutre", "type": "Material"},
    "Iron Pad": {"french": "Plaque de fixation", "type": "Material"},
    "Polyester": {"french": "Polyester", "type": "Material"},
    "Nylon": {"french": "Nylon", "type": "Material"},
    "Acrylic": {"french": "Acide Acroléique", "type": "Material"},
    "Linen": {"french": "Lin", "type": "Material"},
    "Cotton": {"french": "Coton", "type": "Material"},
    "Viscose": {"french": "La fibre artificielle de viscose", "type": "Material"},
    "Recycle polyester": {"french": "Polyester recyclé", "type": "Material"},
    "Wool": {"french": "Laine", "type": "Material"},
    "Polypropylene": {"french": "Polypropylène", "type": "Material"},
    "Olefin": {"french": "Oléfine", "type": "Material"},
    "Polyvinyl Chloride (PVC)": {"french": "Chlorure de Polyvinyle", "type": "Material"},
    "Leather": {"french": "Cuir", "type": "Material"},
    "Seat Cushion Cover": {"french": "Housse de coussin de siege", "type": "Part Title"},
    "Back Cushion Cover": {"french": "Housse de coussin de dossier", "type": "Part Title"},
    "Base Cushion Cover": {"french": "Housse de coussin de base", "type": "Part Title"},
    "Pillow Cover": {"french": "Taie d'oreiller", "type": "Part Title"},
    "Bed Frame Cover": {"french": "Une housse de lit", "type": "Part Title"},
    "Outer Cover": {"french": "Couverture Extérieure", "type": "Part Title"},
    "Mattress Cover": {"french": "Housse de Matelas", "type": "Part Title"},
    "Primary Fabric": {"french": "Tissu Principal", "type": "Sub-part Title"},
    "Secondary Fabric": {"french": "Tissu Secondaire", "type": "Sub-part Title"},
    "Sofa Seat Cushion Cover":{"french": "Housse de coussin d'assise de canapé", "type": "Part Title"},
    "Ice Silk":{"french": "Tissu soie glacée pour matelas", "type": "Material"},
}

# Label 4 distributor blocks by firm (uppercase)
LABEL4_DISTRIBUTOR_LINES = {
    "CASTLERY": [
        "Castlery Private Limited",
        "601 Macpherson Road, Grantral Complex,#07-01 Singapore 368242",
        "Castlery Pty Ltd",
        "1198 Toorak Road, CAMBERWELL VIC 3124 AUSTRALIA",
        "Castlery Inc",
        "1950 W. CORPORATE WAY PMB 95972 ANAHEIM CA 92801, USA",
        "Castlery Ltd.",
        "333 Bay Street, Suite 2400, Toronto, Ontario, M5H 2T6, Canada",
        "Castlery Ltd.",
        "1 GILTSPUR STREET FARRINGDON, LONDON, EC1A 9DD, UNITED KINGDOM",
    ],
    "MOPIO": [
        "Mopio Inc.",
        "5101 Santa Monica Blvd Ste 8 -708 Los Angeles, CA 90029, USA",
        "Mopio Ltd.",
        "333 Bay Street, Suite 2400, Toronto, Ontario, Canada M5H 2T6",
        "Mopio Furniture Limited",
        "1 GILTSPUR STREET FARRINGDON, LONDON, EC1A 9DD, UNITED KINGDOM",
    ],
}

# Label 4 "MADE FOR" company name by firm (uppercase)
LABEL4_MADE_FOR_COMPANY = {
    "CASTLERY": "CASTLERY INC.",
    "MOPIO": "MOPIO INC.",
}

# Label 4 washing instruction mapping (normalized instruction -> icon key)
LABEL4_WASHING_ICON_KEY_BY_INSTRUCTION = {
    "DO NOT BLEACH": "no_bleach",
    "DO NOT DRY CLEAN": "no_dry_clean",
    "DO NOT IRON": "no_iron",
    "DO NOT TUMBLE DRY": "no_tumble_dry",
    "DO NOT WASH WITH WATER": "no_wash",
    "IRONING WITH LOW TEMPERATURE": "iron_low_temp",
    "LINE DRY": "line_dry",
    "SPOT CLEAN": "spot_clean",
    "TUMBLE DRY WITH LOW HEAT": "tumble_dry_low_heat",
    "WASH WITH COLD WATER GENTLY": "wash_cold_gentle",
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


def find_label4_material_match(text: str) -> tuple:
    """
    Backward-compatible wrapper returning only key + French translation.
    For richer metadata, use find_label4_title_match.
    """
    matched_key, french, _line_type = find_label4_title_match(text)
    if matched_key:
        return matched_key, french
    return None, None


def find_label4_title_match(text: str) -> tuple:
    """
    Check if text exactly matches a label4 key, case-insensitive.
    A single trailing colon is tolerated for lookup only.
    Returns (matched_key, french_translation, line_type) or (None, None, None).
    """
    cleaned = text.strip()
    if cleaned.endswith(":"):
        cleaned = cleaned[:-1].strip()
    normalized = cleaned.upper()

    for key, meta in LABEL4_TITLE_MAP.items():
        if key.upper() == normalized:
            return key, meta["french"], meta["type"]
    return None, None, None


def normalize_washing_instruction(text: str) -> str:
    """Normalize washing instruction text for deterministic mapping."""
    import re

    normalized = text.replace("（", "(").replace("）", ")")
    normalized = re.sub(r"\([^)]*\)", "", normalized)
    normalized = normalized.upper()
    normalized = re.sub(r"[^A-Z0-9]+", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def find_label4_washing_icon_key(text: str) -> str:
    """
    Map a washing instruction line to a label4 icon key.
    Returns icon key string or None if no mapping exists.
    """
    normalized = normalize_washing_instruction(text)

    if normalized in LABEL4_WASHING_ICON_KEY_BY_INSTRUCTION:
        return LABEL4_WASHING_ICON_KEY_BY_INSTRUCTION[normalized]

    # Fallback rules for variants with extra words/order differences.
    if "DO NOT" in normalized and "BLEACH" in normalized:
        return "no_bleach"
    if "DO NOT" in normalized and "DRY CLEAN" in normalized:
        return "no_dry_clean"
    if "DO NOT" in normalized and "IRON" in normalized:
        return "no_iron"
    if "DO NOT" in normalized and "TUMBLE" in normalized and "DRY" in normalized:
        return "no_tumble_dry"
    if "DO NOT" in normalized and "WASH" in normalized and "WATER" in normalized:
        return "no_wash"
    if "TUMBLE" in normalized and "LOW HEAT" in normalized:
        return "tumble_dry_low_heat"
    if "IRON" in normalized and "LOW" in normalized:
        return "iron_low_temp"
    if "WASH" in normalized and "COLD" in normalized:
        return "wash_cold_gentle"
    if "LINE DRY" in normalized:
        return "line_dry"
    if "SPOT CLEAN" in normalized:
        return "spot_clean"

    return None


def get_label4_distributor_lines(firm: str) -> list:
    """Get label4 distributor block lines by firm name."""
    normalized_firm = normalize_text(firm) if firm else ""
    if normalized_firm in LABEL4_DISTRIBUTOR_LINES:
        return LABEL4_DISTRIBUTOR_LINES[normalized_firm]

    # Handle common variants from column E (for example: "Mopio Inc.", "Castlery Pte Ltd")
    if "CASTLERY" in normalized_firm:
        return LABEL4_DISTRIBUTOR_LINES.get("CASTLERY", [])
    if "MOPIO" in normalized_firm:
        return LABEL4_DISTRIBUTOR_LINES.get("MOPIO", [])

    return []


def get_label4_made_for_company(firm: str) -> str:
    """Get label4 company name shown under 'MADE FOR' by firm name."""
    normalized_firm = normalize_text(firm) if firm else ""
    if normalized_firm in LABEL4_MADE_FOR_COMPANY:
        return LABEL4_MADE_FOR_COMPANY[normalized_firm]

    if "CASTLERY" in normalized_firm:
        return LABEL4_MADE_FOR_COMPANY.get("CASTLERY", "")
    if "MOPIO" in normalized_firm:
        return LABEL4_MADE_FOR_COMPANY.get("MOPIO", "")

    return ""
