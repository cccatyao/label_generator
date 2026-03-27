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
from generate_label2 import _configure_fontconfig, HAS_CAIROSVG, sanitize_filename

# Import validation functions from centralized module
from validation import contains_non_english_chars, parse_material_text

# Ensure fonts are configured
_configure_fontconfig()

if HAS_CAIROSVG:
    import cairosvg

try:
    from PIL import ImageFont
except Exception:  # pragma: no cover - optional dependency in some runtimes
    ImageFont = None


LABEL19_TEXT_MAX_WIDTH = 358.0
LABEL19_MEASURE_FONT_SIZE = 240
LABEL19_TEXT_FONT_SIZE = 12.19
LABEL19_FONT_PATHS = {
    "cls-5": os.path.join(os.path.dirname(os.path.abspath(__file__)), "font", "AvenirNextCondensed-Medium.ttf"),
    "cls-12": os.path.join(os.path.dirname(os.path.abspath(__file__)), "font", "AvenirNextCondensed-DemiBold.ttf"),
}
LABEL19_FALLBACK_CHAR_WIDTH = {
    "cls-5": 0.34,
    "cls-12": 0.34,
}


def _load_label19_measure_font(css_class: str):
    if ImageFont is None:
        return None
    font_path = LABEL19_FONT_PATHS.get(css_class)
    if not font_path or not os.path.exists(font_path):
        return None
    try:
        return ImageFont.truetype(font_path, LABEL19_MEASURE_FONT_SIZE)
    except Exception:
        return None


def _measure_label19_text_width(text: str, css_class: str) -> float:
    normalized = text.strip()
    if not normalized:
        return 0.0

    font = _load_label19_measure_font(css_class)
    if font is not None:
        try:
            return float(font.getlength(normalized)) * (
                LABEL19_TEXT_FONT_SIZE / LABEL19_MEASURE_FONT_SIZE
            )
        except Exception:
            pass

    return len(normalized) * LABEL19_FALLBACK_CHAR_WIDTH.get(css_class, 0.34) * LABEL19_TEXT_FONT_SIZE


def _split_label19_long_token(token: str, css_class: str, max_width: float) -> List[str]:
    chunks: List[str] = []
    current = ""

    for char in token:
        candidate = f"{current}{char}"
        if current and _measure_label19_text_width(candidate, css_class) > max_width:
            chunks.append(current)
            current = char
        else:
            current = candidate

    if current:
        chunks.append(current)

    return chunks or [token]


def _wrap_label19_line(text: str, css_class: str, max_width: float = LABEL19_TEXT_MAX_WIDTH) -> List[str]:
    normalized = re.sub(r"\s+", " ", text.strip())
    if not normalized:
        return []

    if _measure_label19_text_width(normalized, css_class) <= max_width:
        return [normalized]

    wrapped: List[str] = []
    current = ""

    for token in normalized.split(" "):
        if not token:
            continue

        if _measure_label19_text_width(token, css_class) > max_width:
            if current:
                wrapped.append(current)
                current = ""
            wrapped.extend(_split_label19_long_token(token, css_class, max_width))
            continue

        candidate = token if not current else f"{current} {token}"
        if _measure_label19_text_width(candidate, css_class) <= max_width:
            current = candidate
        else:
            wrapped.append(current)
            current = token

    if current:
        wrapped.append(current)

    return wrapped


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
            for line in _wrap_label19_line(part["title"], TITLE_CLASS):
                tspan_elements.append(
                    f'<tspan class="{TITLE_CLASS}">'
                    f'<tspan x="0" y="{current_y:.2f}" text-anchor="middle">{line}</tspan>'
                    f'</tspan>'
                )
                current_y += LINE_HEIGHT
        
        # Add materials (centered)
        for material in part['materials']:
            for line in _wrap_label19_line(material["text"], MATERIAL_CLASS):
                tspan_elements.append(
                    f'<tspan class="{MATERIAL_CLASS}">'
                    f'<tspan x="0" y="{current_y:.2f}" text-anchor="middle">{line}</tspan>'
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
    
    Note: This function assumes records have already been validated.
    Use validate_record_for_labels() before calling this function.
    
    Uses the same data structure as label 2 - specifically the material_text column.
    
    Args:
        template_content: SVG template content as string
        df: DataFrame with label data (pre-validated)
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
        
        # Skip if missing required fields (shouldn't happen with pre-validation)
        if not materials_text:
            continue
        
        # Parse material text into parts (validation already done, so this should succeed)
        parts, parse_alerts, has_unmapped_terms = parse_material_text(materials_text)
        
        if not parts:
            # This shouldn't happen with pre-validation, but keep as safety check
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
