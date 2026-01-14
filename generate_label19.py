#!/usr/bin/env python3
"""
Label 19 Generator - Filling Label
Generates SVG and PDF filling labels from template and Excel data.

Label 19 displays the material composition organized by part, with:
- Part titles (e.g., "SEAT CUSHION(French text)(qty):")
- Composition materials listed under each part (must start with percentage)
- Visual separation between different parts
"""

import os
import re
import pandas as pd
from typing import List, Tuple, Optional, Dict

# Configure fontconfig (reuse from generate_label2)
from generate_label2 import _configure_fontconfig, HAS_CAIROSVG, sanitize_filename, contains_non_english_chars

# Ensure fonts are configured
_configure_fontconfig()

if HAS_CAIROSVG:
    import cairosvg


def parse_material_text(material_text: str) -> Tuple[List[Dict], List[str]]:
    """
    Parse material_text into structured parts using dictionary-based detection.
    
    Args:
        material_text: Multi-line text with parts and materials
        
    Returns:
        Tuple of (parts_list, alerts) where:
        - parts_list: List of dicts with 'title' and 'materials' keys
        - alerts: List of alert messages for validation issues
    """
    from term_config import find_part_match, find_material_match, normalize_text
    
    # Normalize line breaks
    text = material_text.replace('\\n', '\n')
    lines = text.split('\n')
    
    parts = []
    alerts = []
    current_part = None
    
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
            # Not found in dictionary - add warning
            alerts.append(f"part or material not exist in dictionary: '{line}'")
            # Still process it as best effort - treat as part if no current part, else as material
            if current_part is None:
                if current_part is not None and (current_part['title'] or current_part['materials']):
                    parts.append(current_part)
                current_part = {
                    'title': f"{line}:",
                    'materials': []
                }
            else:
                current_part['materials'].append({
                    'text': line,
                    'has_percentage': False
                })
    
    # Don't forget the last part
    if current_part is not None and (current_part['title'] or current_part['materials']):
        parts.append(current_part)
    
    return parts, alerts


def generate_label19_svg_content(parts: List[Dict]) -> Tuple[str, float]:
    """
    Generate the dynamic text content for label 19 SVG.
    
    Args:
        parts: Parsed parts with titles and materials
        
    Returns:
        Tuple of (SVG tspan elements, total content height)
    """
    # Text styling parameters (from template analysis)
    TITLE_CLASS = "cls-12"  # DemiBold for titles
    MATERIAL_CLASS = "cls-5"  # Medium for materials
    LINE_HEIGHT = 14.63
    PART_SPACING = 29.26  # Double line height for part separation
    
    tspan_elements = []
    current_y = 0
    
    for part_idx, part in enumerate(parts):
        # Add extra spacing before parts (except first one)
        if part_idx > 0:
            current_y += PART_SPACING - LINE_HEIGHT  # Additional gap
        
        # Add part title (centered)
        if part['title']:
            tspan_elements.append(
                f'<tspan class="{TITLE_CLASS}">'
                f'<tspan x="0" y="{current_y:.2f}" text-anchor="middle">{part["title"]}</tspan>'
                f'</tspan>'
            )
            current_y += LINE_HEIGHT
        
        # Add materials (centered)
        for material in part['materials']:
            tspan_elements.append(
                f'<tspan class="{MATERIAL_CLASS}">'
                f'<tspan x="0" y="{current_y:.2f}" text-anchor="middle">{material["text"]}</tspan>'
                f'</tspan>'
            )
            current_y += LINE_HEIGHT
    
    return ''.join(tspan_elements), current_y


def replace_label19_variables(svg_content: str, parts_content: str, content_height: float) -> str:
    """
    Replace template variables in the label 19 SVG content.
    
    Args:
        svg_content: Original SVG template content
        parts_content: Generated tspan elements for parts
        content_height: Total height of the text content (Y offset after last line)
        
    Returns:
        Modified SVG content
    """
    # Constants from template
    LABEL_TOP = 23.2  # Top of the label rectangle
    TEXT_START_Y = 80.24  # Y position where dynamic text starts
    REMPLISSAGE_Y = 58.42  # Y position of "(Remplissage)" text
    LINE_HEIGHT = 14.63
    FONT_SIZE = 12.19  # Font size from cls-9
    
    # Spacing calculation: In the SVG, 360px = 127mm (from the width dimension line)
    # Therefore: 1mm = 360/127 = 2.8346px
    # Required spacing: 3.84mm = 3.84 × 2.8346 = 10.88px
    VISUAL_SPACING = 10.88  # 3.84mm in pixels
    
    # The content_height represents the y-offset after the last line
    # We need to subtract one LINE_HEIGHT because we don't want the spacing after the last line
    # The actual visual content ends at the last line's baseline
    actual_content_height = content_height - LINE_HEIGHT
    
    # Calculate label dimensions with equal visual top and bottom spacing
    BOTTOM_SPACING = VISUAL_SPACING  # Same visual spacing at bottom as top
    
    # Total label height: from label top to last text line + bottom spacing
    # = (distance from label top to text start) + actual_content_height + bottom_spacing
    label_height = (TEXT_START_Y - LABEL_TOP) + actual_content_height + BOTTOM_SPACING
    label_bottom = LABEL_TOP + label_height
    
    # Replace all placeholders
    svg_content = svg_content.replace('{{parts_content}}', parts_content)
    svg_content = svg_content.replace('{{label_height}}', f'{label_height:.2f}')
    svg_content = svg_content.replace('{{label_bottom}}', f'{label_bottom:.2f}')
    
    return svg_content


def convert_svg_to_pdf(svg_content: str) -> Optional[bytes]:
    """Convert SVG content to PDF bytes."""
    if not HAS_CAIROSVG:
        return None
    try:
        pdf_bytes = cairosvg.svg2pdf(bytestring=svg_content.encode('utf-8'))
        return pdf_bytes
    except Exception as e:
        print(f"PDF conversion failed: {e}")
        return None


def generate_label19_from_dataframe(
    template_content: str,
    df: pd.DataFrame,
    generate_pdf: bool = True
) -> Tuple[List[Tuple[str, bytes]], List[str]]:
    """
    Generate Label 19 PDFs from a DataFrame (in-memory, no file I/O).
    
    Uses the same data structure as label 2 - specifically the material_text column.
    
    Args:
        template_content: SVG template content as string
        df: DataFrame with label data (same format as label 2)
        generate_pdf: Whether to generate PDF files
        
    Returns:
        Tuple of (pdf_files, warnings) where:
        - pdf_files: list of (filename, content) tuples
        - warnings: list of warning messages
    """
    columns = df.columns.tolist()
    materials_col = columns[1]  # Material composition text
    
    pdf_files = []
    warnings = []
    
    for index, row in df.iterrows():
        materials_text = str(row[materials_col]) if pd.notna(row[materials_col]) else ""
        identifier = str(row[columns[0]]) if pd.notna(row[columns[0]]) else f"label_{index}"
        
        if not materials_text:
            continue
        
        # Validate English input
        if contains_non_english_chars(materials_text):
            warnings.append(f"{identifier} label19 is not generated, reason: material text is not English input.")
            continue
        
        # Parse material text into parts
        parts, parse_alerts = parse_material_text(materials_text)
        
        # Add parsing alerts as warnings
        for alert in parse_alerts:
            warnings.append(f"{identifier} label19: {alert}")
        
        if not parts:
            warnings.append(f"{identifier} label19 is not generated, reason: no valid parts found in material text.")
            continue
        
        # Generate SVG content
        parts_content, content_height = generate_label19_svg_content(parts)
        svg_content = replace_label19_variables(template_content, parts_content, content_height)
        
        # Generate PDF
        safe_name = sanitize_filename(identifier)
        pdf_filename = f"{safe_name}-label19.pdf"
        
        if HAS_CAIROSVG and generate_pdf:
            pdf_bytes = convert_svg_to_pdf(svg_content)
            if pdf_bytes:
                pdf_files.append((pdf_filename, pdf_bytes))
    
    return pdf_files, warnings
