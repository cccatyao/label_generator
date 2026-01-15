#!/usr/bin/env python3
"""
Validation Module for Label Generator
Contains all validation logic for both Label 2 and Label 19.
"""

import re
from typing import List, Tuple, Dict


def contains_non_english_chars(text: str) -> bool:
    """
    Check if text contains non-English characters (like Chinese parentheses).
    Returns True if non-English characters are found.
    """
    # Common non-English characters to check for
    non_english_chars = [
        '（', '）',  # Chinese parentheses
        '【', '】',  # Chinese brackets
        '「', '」',  # Chinese quotation marks
        '『', '』',  # Double angle brackets
        '《', '》',  # Chinese book title marks
        '，', '。',  # Chinese comma and period
        '：', '；',  # Chinese colon and semicolon
        '"', '"',   # Chinese quotation marks
        ''', ''',   # Chinese single quotes
        '、',       # Chinese enumeration comma
        '％',       # Full-width percent
    ]
    
    for char in non_english_chars:
        if char in text:
            return True
    
    # Also check for characters outside basic ASCII printable range (except common unicode)
    for char in text:
        # Allow ASCII printable characters, newlines, and some common symbols
        if ord(char) > 127:
            # Check if it's a common acceptable unicode (like degree symbol, etc.)
            # For now, flag any non-ASCII as potentially non-English
            if char not in ['°', '±', '×', '÷', '®', '™', '©']:
                return True
    
    return False


def parse_material_text(material_text: str) -> Tuple[List[Dict], List[str], bool]:
    """
    Parse material_text into structured parts using dictionary-based detection.
    
    Args:
        material_text: Multi-line text with parts and materials
        
    Returns:
        Tuple of (parts_list, alerts, has_unmapped_terms) where:
        - parts_list: List of dicts with 'title' and 'materials' keys
        - alerts: List of alert messages for validation issues
        - has_unmapped_terms: True if any terms couldn't be mapped to dictionary
    """
    from term_config import find_part_match, find_material_match
    
    # Normalize line breaks
    text = material_text.replace('\\n', '\n')
    lines = text.split('\n')
    
    parts = []
    alerts = []
    current_part = None
    has_unmapped_terms = False
    
    # Pattern to extract: text before count (1), (2), etc.
    # Matches optional count like (1), (2) at the end
    count_pattern = re.compile(r'^(.*?)(\(\d+\))?\s*:?\s*$')
    # Pattern to detect percentage at start of line (e.g., "98%", "2%", "100%")
    percentage_pattern = re.compile(r'^(\d+%)\s*(.*)$')
    
    for line in lines:
        line = line.strip()
        
        if not line:
            continue
        
        # Check if line is a Part or Material using dictionary
        part_key, part_french = find_part_match(line)
        material_key, material_french = find_material_match(line)
        
        if part_key:
            # This is a part title
            # Save previous part if exists
            if current_part is not None and (current_part['title'] or current_part['materials']):
                parts.append(current_part)
            
            # Process the title: extract count suffix if present
            match = count_pattern.match(line)
            if match:
                title_text = match.group(1).strip()
                count_suffix = match.group(2) or ""
            else:
                title_text = line.rstrip(':').strip()
                count_suffix = ""
            
            # Remove parentheses from English title (e.g., BODY(ARMLESS) -> BODY ARMLESS)
            title_text = re.sub(r'\(([^)]+)\)', r' \1', title_text)
            title_text = re.sub(r'\s+', ' ', title_text).strip()
            
            # Build final title with French: {English}({French}){count}:
            formatted_title = f"{title_text}({part_french}){count_suffix}:"
            
            current_part = {
                'title': formatted_title,
                'materials': []
            }
        elif material_key:
            # This is a material line
            if current_part is None:
                current_part = {
                    'title': None,
                    'materials': []
                }
            
            # Extract percentage if present
            pct_match = percentage_pattern.match(line)
            if pct_match:
                percentage = pct_match.group(1)
                material_text_part = pct_match.group(2).strip()
                # Build formatted material: {percentage} {English}({French})
                formatted_material = f"{percentage} {material_text_part}({material_french})"
            else:
                # No percentage, just add French translation
                formatted_material = f"{line}({material_french})"
            
            current_part['materials'].append({
                'text': formatted_material,
                'has_percentage': bool(pct_match)
            })
        else:
            # Not found in dictionary - add warning and mark as unmapped
            alerts.append(f"part or material not exist in dictionary: '{line}'")
            has_unmapped_terms = True
            # Don't process unmapped terms - skip label generation for this record
    
    # Don't forget the last part
    if current_part is not None and (current_part['title'] or current_part['materials']):
        parts.append(current_part)
    
    return parts, alerts, has_unmapped_terms


def validate_record_for_labels(
    materials_text: str,
    reg_no: str,
    per_no: str,
    identifier: str
) -> Tuple[bool, List[str]]:
    """
    Validate a record against combined rules for both Label 2 and Label 19.
    
    Validation Rules:
    1. Material text must not be empty
    2. REG number must not be empty
    3. Material text must be ≤ 15 lines
    4. Material text must be English (no non-English characters)
    5. REG number must be English
    6. PER number (if present) must be English
    7. All parts/materials must exist in term_config.py dictionary
    
    Args:
        materials_text: Material composition text
        reg_no: Registration number
        per_no: PER number (optional)
        identifier: Record identifier for error messages
        
    Returns:
        Tuple of (is_valid, error_messages) where:
        - is_valid: True if record passes all validations
        - error_messages: List of validation error messages
    """
    errors = []
    MAX_MATERIAL_LINES = 15
    
    # Rule 1 & 2: Check if required fields are present
    if not materials_text or not reg_no:
        if not materials_text:
            errors.append(f"{identifier} labels are not generated, reason: material text is empty.")
        if not reg_no:
            errors.append(f"{identifier} labels are not generated, reason: REG number is empty.")
        return False, errors
    
    # Rule 3: Check material text line count
    material_lines = materials_text.replace('\\n', '\n').split('\n')
    non_empty_lines = [line for line in material_lines if line.strip()]
    if len(non_empty_lines) > MAX_MATERIAL_LINES:
        errors.append(f"{identifier} labels are not generated, reason: material text larger than {MAX_MATERIAL_LINES} lines.")
        return False, errors
    
    # Rule 4: Validate English input for material_text
    if contains_non_english_chars(materials_text):
        errors.append(f"{identifier} labels are not generated, reason: material text is not English input.")
        return False, errors
    
    # Rule 5: Validate English input for reg_no
    if contains_non_english_chars(reg_no):
        errors.append(f"{identifier} labels are not generated, reason: REG # is not English input.")
        return False, errors
    
    # Rule 6: Validate English input for per_no (if present)
    if per_no and contains_non_english_chars(per_no):
        errors.append(f"{identifier} labels are not generated, reason: PER # is not English input.")
        return False, errors
    
    # Rule 7: Validate that all parts/materials are in dictionary (for Label 19)
    parts, parse_alerts, has_unmapped_terms = parse_material_text(materials_text)
    
    if has_unmapped_terms:
        # Add specific errors about unmapped terms
        for alert in parse_alerts:
            errors.append(f"{identifier} labels are not generated, reason: {alert}")
        return False, errors
    
    # All validations passed
    return True, []
