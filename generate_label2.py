#!/usr/bin/env python3
"""
Law Label Generator - Core Functions
Generates SVG and PDF law labels from template and Excel data.

This module contains the core label generation functions that can be called
from both CLI and Streamlit web interface.
"""

import os
import re
import io
import platform
from functools import lru_cache
import pandas as pd
from typing import List, Tuple, Optional

# Import validation functions from centralized module
from validation import contains_non_english_chars
from term_config import ORIGIN_COUNTRY_MAP


def _configure_cairo_library_path():
    """Help cairocffi find Homebrew cairo libraries on macOS."""
    if platform.system() != "Darwin":
        return

    existing = [
        path
        for path in os.environ.get("DYLD_FALLBACK_LIBRARY_PATH", "").split(":")
        if path
    ]

    preferred_dirs = []
    if platform.machine().lower() == "arm64":
        preferred_dirs.append("/opt/homebrew/lib")
    elif platform.machine().lower() in {"x86_64", "amd64"}:
        preferred_dirs.append("/usr/local/lib")

    for candidate in ("/opt/homebrew/lib", "/usr/local/lib"):
        if candidate not in preferred_dirs:
            preferred_dirs.append(candidate)

    updated = [
        candidate
        for candidate in preferred_dirs
        if os.path.isdir(candidate) and candidate not in existing
    ] + existing

    if updated:
        os.environ["DYLD_FALLBACK_LIBRARY_PATH"] = ":".join(updated)


# Configure fontconfig to use local font folder before importing cairosvg
# This ensures the Avenir Next Condensed font is available even when not installed system-wide
def _configure_fontconfig():
    """Configure fontconfig to include the project's font directory."""
    import subprocess
    import tempfile
    
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    font_dir = os.path.join(script_dir, 'font')
    
    if not os.path.exists(font_dir):
        return
    
    # Create a comprehensive fonts.conf that includes both custom and system fonts
    fonts_conf_content = f'''<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd">
<fontconfig>
    <!-- Include default system configuration -->
    <include ignore_missing="yes">/etc/fonts/fonts.conf</include>
    <include ignore_missing="yes">/etc/fonts/conf.d</include>
    <include ignore_missing="yes">/usr/share/fonts</include>
    <include ignore_missing="yes">/usr/local/share/fonts</include>
    
    <!-- Add project's custom font directory -->
    <dir>{font_dir}</dir>
    
    <!-- Cache directory -->
    <cachedir prefix="xdg">fontconfig</cachedir>
    <cachedir>/tmp/fontconfig-cache</cachedir>
    
    <!-- Font matching rules -->
    <match target="pattern">
        <test name="family" qual="any">
            <string>AvenirNextCondensed-Bold</string>
        </test>
        <edit name="family" mode="assign" binding="strong">
            <string>Avenir Next Condensed Bold</string>
        </edit>
    </match>
    
    <match target="pattern">
        <test name="family" qual="any">
            <string>AvenirNextCondensed-Medium</string>
        </test>
        <edit name="family" mode="assign" binding="strong">
            <string>Avenir Next Condensed Medium</string>
        </edit>
    </match>
    
    <match target="pattern">
        <test name="family" qual="any">
            <string>AvenirNextCondensed-DemiBold</string>
        </test>
        <edit name="family" mode="assign" binding="strong">
            <string>Avenir Next Condensed Demi Bold</string>
        </edit>
    </match>
    
    <match target="pattern">
        <test name="family" qual="any">
            <string>AvenirNextCondensed-UltraLight</string>
        </test>
        <edit name="family" mode="assign" binding="strong">
            <string>Avenir Next Condensed Ultra Light</string>
        </edit>
    </match>
</fontconfig>
'''
    
    fonts_conf_path = os.path.join(font_dir, 'fonts.conf')
    try:
        with open(fonts_conf_path, 'w') as f:
            f.write(fonts_conf_content)
    except Exception:
        pass  # Ignore if we can't write the file
    
    # Set fontconfig environment variables
    os.environ['FONTCONFIG_FILE'] = fonts_conf_path
    os.environ['FONTCONFIG_PATH'] = font_dir
    
    # Force fontconfig cache rebuild to pick up new fonts
    try:
        cache_dir = '/tmp/fontconfig-cache'
        os.makedirs(cache_dir, exist_ok=True)
        # Run fc-cache to build font cache (silently)
        subprocess.run(['fc-cache', '-f', font_dir], 
                      capture_output=True, timeout=30)
    except Exception:
        pass  # fc-cache may not be available

# Configure native library and fonts before importing cairosvg
_configure_cairo_library_path()
_configure_fontconfig()

# Try to import cairosvg for PDF conversion
try:
    import cairosvg
    HAS_CAIROSVG = True
except (ImportError, OSError):
    HAS_CAIROSVG = False

try:
    from PIL import ImageFont
except Exception:  # pragma: no cover - optional dependency in some runtimes
    ImageFont = None


LABEL2_INNER_WIDTH = 182.37
LABEL2_TEXT_MAX_WIDTH = LABEL2_INNER_WIDTH - 2.0
LABEL2_MEASURE_FONT_SIZE = 240
LABEL2_TEXT_FONT_SIZE = 13.32
LABEL2_FALLBACK_CHAR_WIDTH = 0.34
LABEL2_LINE_HEIGHT = 15.99
LABEL2_BASE_SVG_HEIGHT = 841.89
LABEL2_BASE_LABEL_HEIGHT = 568.69
LABEL2_BASE_LABEL_BOTTOM_Y = 593.99
LABEL2_MATERIAL_START_Y = 124.6
LABEL2_MATERIAL_END_Y = 353.08
LABEL2_MATERIAL_TOP_PADDING = 6.0
LABEL2_MATERIAL_BOTTOM_PADDING = 10.0
LABEL2_CODE_SECTION_TOP_Y = 353.08
LABEL2_CODE_SECTION_BOTTOM_Y = 397.83
LABEL2_CODE_SECTION_HEIGHT = round(
    LABEL2_CODE_SECTION_BOTTOM_Y - LABEL2_CODE_SECTION_TOP_Y,
    2,
)
LABEL2_MEASURE_FONT_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "font",
    "AvenirNextCondensed-DemiBold.ttf",
)


@lru_cache(maxsize=1)
def _load_label2_measure_font():
    if ImageFont is None or not os.path.exists(LABEL2_MEASURE_FONT_PATH):
        return None
    try:
        return ImageFont.truetype(LABEL2_MEASURE_FONT_PATH, LABEL2_MEASURE_FONT_SIZE)
    except Exception:
        return None


def _measure_label2_text_width(text: str) -> float:
    normalized = text.strip()
    if not normalized:
        return 0.0

    font = _load_label2_measure_font()
    if font is not None:
        try:
            return float(font.getlength(normalized)) * (
                LABEL2_TEXT_FONT_SIZE / LABEL2_MEASURE_FONT_SIZE
            )
        except Exception:
            pass

    return len(normalized) * LABEL2_FALLBACK_CHAR_WIDTH * LABEL2_TEXT_FONT_SIZE


def _split_label2_long_token(token: str, max_width: float) -> List[str]:
    chunks: List[str] = []
    current = ""

    for char in token:
        candidate = f"{current}{char}"
        if current and _measure_label2_text_width(candidate) > max_width:
            chunks.append(current)
            current = char
        else:
            current = candidate

    if current:
        chunks.append(current)

    return chunks or [token]


def _wrap_label2_line(text: str, max_width: float = LABEL2_TEXT_MAX_WIDTH) -> List[str]:
    normalized = re.sub(r"\s+", " ", text.strip())
    if not normalized:
        return []

    if _measure_label2_text_width(normalized) <= max_width:
        return [normalized]

    wrapped: List[str] = []
    current = ""

    for token in normalized.split(" "):
        if not token:
            continue

        if _measure_label2_text_width(token) > max_width:
            if current:
                wrapped.append(current)
                current = ""
            wrapped.extend(_split_label2_long_token(token, max_width))
            continue

        candidate = token if not current else f"{current} {token}"
        if _measure_label2_text_width(candidate) <= max_width:
            current = candidate
        else:
            wrapped.append(current)
            current = token

    if current:
        wrapped.append(current)

    return wrapped


def _wrap_label2_text_lines(text: str) -> List[str]:
    wrapped_lines: List[str] = []
    for line in text.replace("\\n", "\n").split("\n"):
        if not line.strip():
            wrapped_lines.append("")
            continue
        wrapped_lines.extend(_wrap_label2_line(line))
    return wrapped_lines


def create_centered_tspan_elements(
    text: str,
    line_height: float = LABEL2_LINE_HEIGHT,
    top_padding: float = LABEL2_MATERIAL_TOP_PADDING,
) -> str:
    """
    Create tspan elements from multi-line text with each line horizontally centered.
    
    Args:
        text: Multi-line text to convert (can use \\n or actual newlines)
        line_height: Height between lines
        
    Returns:
        String containing tspan elements
    """
    lines = _wrap_label2_text_lines(text)
    
    tspan_elements = []
    current_y = top_padding
    
    for i, line in enumerate(lines):
        line_content = line.strip()
        
        if not line_content:
            current_y += line_height
            continue
        
        if i == 0:
            tspan = f'<tspan x="0" y="{current_y}">{line_content}</tspan>'
        else:
            tspan = f'<tspan x="0" y="{current_y:.2f}">{line_content}</tspan>'
        
        tspan_elements.append(tspan)
        current_y += line_height
    
    return ''.join(tspan_elements)


def _calculate_label2_layout_offset(
    material_text: str,
    line_height: float = LABEL2_LINE_HEIGHT,
) -> float:
    wrapped_line_count = len(_wrap_label2_text_lines(material_text))
    if wrapped_line_count <= 1:
        return 0.0

    available_height = LABEL2_MATERIAL_END_Y - LABEL2_MATERIAL_START_Y
    used_height = (
        LABEL2_MATERIAL_TOP_PADDING
        + (wrapped_line_count - 1) * line_height
        + LABEL2_MATERIAL_BOTTOM_PADDING
    )
    return round(max(0.0, used_height - available_height), 2)


def _wrap_label2_fragment(
    svg_content: str,
    start_marker: str,
    end_marker: str,
    offset: float,
) -> str:
    if offset == 0:
        return svg_content

    start_index = svg_content.index(start_marker)
    end_index = svg_content.index(end_marker, start_index)
    fragment = svg_content[start_index:end_index]
    wrapped_fragment = (
        f'   <g transform="translate(0 {offset:.2f})">\n'
        f"{fragment}"
        "   </g>\n"
    )
    return svg_content[:start_index] + wrapped_fragment + svg_content[end_index:]


def _collapse_label2_code_section(svg_content: str) -> str:
    code_block_start = (
        '   <text class="cls-63" transform="translate(536.5 375.45)" '
        'text-anchor="middle" dominant-baseline="middle"\n'
        '      id="text378">\n'
    )
    code_block_end = '   </text>\n'
    start_index = svg_content.index(code_block_start)
    end_index = svg_content.index(code_block_end, start_index) + len(code_block_end)
    svg_content = svg_content[:start_index] + svg_content[end_index:]

    svg_content = svg_content.replace(
        '   <line class="cls-17" x1="445.21" y1="353.08" x2="627.58" y2="353.08" id="line100" />\n',
        '',
        1,
    )
    return svg_content


def _resize_label2_canvas(svg_content: str, offset: float) -> str:
    if offset <= 0:
        return svg_content

    svg_height = f"{LABEL2_BASE_SVG_HEIGHT + offset:.2f}"
    label_height = f"{LABEL2_BASE_LABEL_HEIGHT + offset:.2f}"
    label_bottom_y = f"{LABEL2_BASE_LABEL_BOTTOM_Y + offset:.2f}"

    svg_content = svg_content.replace(
        f'viewBox="0 0 988.11 {LABEL2_BASE_SVG_HEIGHT:.2f}"',
        f'viewBox="0 0 988.11 {svg_height}"',
        1,
    )
    svg_content = svg_content.replace(
        f'height="{LABEL2_BASE_SVG_HEIGHT:.2f}"',
        f'height="{svg_height}"',
    )
    svg_content = svg_content.replace(
        f'y2="{LABEL2_BASE_LABEL_BOTTOM_Y:.2f}"',
        f'y2="{label_bottom_y}"',
        2,
    )
    svg_content = svg_content.replace(
        f'width="360.06" height="{LABEL2_BASE_LABEL_HEIGHT:.2f}" id="rect222"',
        f'width="360.06" height="{label_height}" id="rect222"',
        1,
    )
    return svg_content


def _extend_label2_layout(svg_content: str, offset: float) -> str:
    if offset <= 0:
        return svg_content

    svg_content = _resize_label2_canvas(svg_content, offset)
    svg_content = _wrap_label2_fragment(
        svg_content,
        '<line class="cls-17" x1="445.21" y1="353.08" x2="627.58" y2="353.08" id="line100" />',
        '<line class="cls-17" x1="445.07" y1="75.11" x2="628.85" y2="75.11" id="line101" />',
        offset,
    )
    svg_content = _wrap_label2_fragment(
        svg_content,
        '<line class="cls-17" x1="445.01" y1="397.83" x2="627.92" y2="397.83" id="line102" />',
        '<line class="cls-20" x1="626.8" y1="260.47" x2="805.25" y2="260.47" id="line219" />',
        offset,
    )
    svg_content = _wrap_label2_fragment(
        svg_content,
        '<text class="cls-63" transform="translate(536.5 375.45)" text-anchor="middle" dominant-baseline="middle"',
        '<rect class="cls-18" x="392.83" width="595.28"',
        offset,
    )
    svg_content = _wrap_label2_fragment(
        svg_content,
        '<text class="cls-62" transform="translate(651.2 378.42)" id="text387">',
        '<text class="cls-63" transform="translate(633.82 277.73)" id="text646">',
        offset,
    )
    svg_content = _wrap_label2_fragment(
        svg_content,
        '<text class="cls-63" transform="translate(508.49 479.27)" id="text743">',
        '</svg>',
        offset,
    )
    return svg_content


def replace_template_variables(svg_content: str, material_text: str, reg_number: str, per_number: str = "", firm: str = "", origin: str = "") -> str:
    """
    Replace template variables in the SVG content.
    
    Args:
        svg_content: Original SVG content
        material_text: Multi-line material composition text
        reg_number: Registration number (without REG.NO. prefix)
        per_number: Optional PER number (without PER.NO. prefix)
        firm: Firm name
        origin: Origin country code (CN or VN)
    """
    reg_number_clean = reg_number.strip() if reg_number else ""
    per_number_clean = per_number.strip() if per_number else ""

    code_number_lines = []
    if reg_number_clean:
        code_number_lines.append(f"REG.NO.{reg_number_clean}")
    if per_number_clean:
        code_number_lines.append(f"PER.NO.{per_number_clean}")

    if len(code_number_lines) == 2:
        code_number_content = (
            f'<tspan x="0" dy="-8">{code_number_lines[0]}</tspan>'
            f'<tspan x="0" dy="16">{code_number_lines[1]}</tspan>'
        )
    elif len(code_number_lines) == 1:
        code_number_content = code_number_lines[0]
    else:
        code_number_content = ""

    svg_content = svg_content.replace('{{code_number}}', code_number_content)
    
    material_tspans = create_centered_tspan_elements(material_text, line_height=LABEL2_LINE_HEIGHT)
    svg_content = svg_content.replace('{{material_text}}', material_tspans)
    
    # Handle firm name
    svg_content = svg_content.replace('{{firm}}', firm.strip() if firm else '')
    
    # Handle origin country - map CN to CHINA, VN to VIETNAM
    origin_clean = origin.strip().upper() if origin else ""
    origin_country = ORIGIN_COUNTRY_MAP.get(origin_clean, origin_clean)
    svg_content = svg_content.replace('{{origin_country}}', origin_country)

    layout_offset = _calculate_label2_layout_offset(material_text)
    svg_content = _extend_label2_layout(svg_content, layout_offset)
    if not code_number_lines:
        svg_content = _collapse_label2_code_section(svg_content)
    return svg_content


def sanitize_filename(text: str) -> str:
    """Create a safe filename from text."""
    safe = re.sub(r'[<>:"/\\|?*\n\r]', '', text)
    safe = safe.replace(' ', '_')
    safe = safe[:50]
    return safe


def convert_svg_bytes_to_pdf_bytes(svg_content: str) -> Optional[bytes]:
    """Convert SVG content to PDF bytes in memory."""
    if not HAS_CAIROSVG:
        return None
    try:
        pdf_bytes = cairosvg.svg2pdf(bytestring=svg_content.encode('utf-8'))
        return pdf_bytes
    except Exception as e:
        print(f"PDF conversion failed: {e}")
        return None


def generate_label2_from_dataframe(
    template_content: str, 
    df: pd.DataFrame, 
    generate_pdf: bool = True
) -> Tuple[List[Tuple[str, bytes]], List[str]]:
    """
    Generate PDF labels from a DataFrame (in-memory, no file I/O).
    
    Note: This function assumes records have already been validated.
    Use validate_record_for_labels() before calling this function.
    
    Args:
        template_content: SVG template content as string
        df: DataFrame with label data (pre-validated)
        generate_pdf: Whether to generate PDF files (kept for compatibility)
        
    Returns:
        Tuple of (pdf_files, warnings) where:
        - pdf_files: list of (filename, content) tuples
        - warnings: list of warning messages for skipped entries
    """
    columns = df.columns.tolist()
    materials_col = columns[1]
    reg_no_col = columns[2]
    # PER. No column is optional (4th column, index 3)
    per_no_col = columns[3] if len(columns) > 3 else None
    # Firm column (5th column, index 4)
    firm_col = columns[4] if len(columns) > 4 else None
    # Origin column (6th column, index 5)
    origin_col = columns[5] if len(columns) > 5 else None
    
    pdf_files = []
    warnings = []
    
    for index, row in df.iterrows():
        materials_text = str(row[materials_col]) if pd.notna(row[materials_col]) else ""
        reg_no = str(row[reg_no_col]) if pd.notna(row[reg_no_col]) else ""
        identifier = str(row[columns[0]]) if pd.notna(row[columns[0]]) else f"label_{index}"
        
        # Get PER. No if column exists
        per_no = ""
        if per_no_col and per_no_col in row:
            per_no = str(row[per_no_col]) if pd.notna(row[per_no_col]) else ""
        
        # Get Firm if column exists
        firm = ""
        if firm_col and firm_col in row:
            firm = str(row[firm_col]) if pd.notna(row[firm_col]) else ""
        
        # Get Origin if column exists
        origin = ""
        if origin_col and origin_col in row:
            origin = str(row[origin_col]) if pd.notna(row[origin_col]) else ""
        
        # Skip if material text is missing (shouldn't happen with pre-validation)
        if not materials_text:
            continue
        
        svg_content = replace_template_variables(template_content, materials_text, reg_no, per_no, firm, origin)
        
        # Generate PDF with new naming pattern: {default_code}-label2.pdf
        safe_name = sanitize_filename(identifier)
        pdf_filename = f"{safe_name}-label2.pdf"
        
        if HAS_CAIROSVG:
            pdf_bytes = convert_svg_bytes_to_pdf_bytes(svg_content)
            if pdf_bytes:
                pdf_files.append((pdf_filename, pdf_bytes))
    
    return pdf_files, warnings
