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
import pandas as pd
from typing import List, Tuple, Optional

# Try to import cairosvg for PDF conversion
try:
    import cairosvg
    HAS_CAIROSVG = True
except ImportError:
    HAS_CAIROSVG = False


def create_centered_tspan_elements(text: str, line_height: float = 15.99) -> str:
    """
    Create tspan elements from multi-line text with each line horizontally centered.
    
    Args:
        text: Multi-line text to convert (can use \\n or actual newlines)
        line_height: Height between lines
        
    Returns:
        String containing tspan elements
    """
    lines = text.replace('\\n', '\n').split('\n')
    
    tspan_elements = []
    current_y = 0
    
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
    # Handle code_number (REG + optional PER)
    formatted_reg_no = f"REG.NO.{reg_number}"
    
    # Check if per_number has a valid value (not empty, not just spaces)
    per_number_clean = per_number.strip() if per_number else ""
    
    if per_number_clean:
        # Two rows: REG.NO. on first line, PER.NO. on second line
        # Use tspan elements with y offsets to center both lines as a whole
        # Line height approximately 16px, so offset each line by half to center
        formatted_per_no = f"PER.NO.{per_number_clean}"
        code_number_content = f'<tspan x="0" dy="-8">{formatted_reg_no}</tspan><tspan x="0" dy="16">{formatted_per_no}</tspan>'
    else:
        # Single row: just REG.NO.
        code_number_content = formatted_reg_no
    
    svg_content = svg_content.replace('{{code_number}}', code_number_content)
    
    material_tspans = create_centered_tspan_elements(material_text, line_height=15.99)
    svg_content = svg_content.replace('{{material_text}}', material_tspans)
    
    # Handle firm name
    svg_content = svg_content.replace('{{firm}}', firm.strip() if firm else '')
    
    # Handle origin country - map CN to CHINA, VN to VIETNAM
    origin_clean = origin.strip().upper() if origin else ""
    origin_map = {'CN': 'CHINA', 'VN': 'VIETNAM'}
    origin_country = origin_map.get(origin_clean, origin_clean)
    svg_content = svg_content.replace('{{origin_country}}', origin_country)
    
    return svg_content


def sanitize_filename(text: str) -> str:
    """Create a safe filename from text."""
    safe = re.sub(r'[<>:"/\\|?*\n\r]', '', text)
    safe = safe.replace(' ', '_')
    safe = safe[:50]
    return safe


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


def generate_labels_from_dataframe(
    template_content: str, 
    df: pd.DataFrame, 
    generate_pdf: bool = True
) -> Tuple[List[Tuple[str, bytes]], List[str]]:
    """
    Generate PDF labels from a DataFrame (in-memory, no file I/O).
    
    Args:
        template_content: SVG template content as string
        df: DataFrame with label data
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
    
    MAX_MATERIAL_LINES = 15
    
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
        
        if not materials_text or not reg_no:
            continue
        
        # Check material text line count (handle both \\n and actual newlines)
        material_lines = materials_text.replace('\\n', '\n').split('\n')
        # Count non-empty lines
        non_empty_lines = [line for line in material_lines if line.strip()]
        if len(non_empty_lines) > MAX_MATERIAL_LINES:
            warnings.append(f"{identifier} label is not generated, reason: material text larger than {MAX_MATERIAL_LINES} lines.")
            continue
        
        # Validate English input for material_text, reg_no, and per_no
        if contains_non_english_chars(materials_text):
            warnings.append(f"{identifier} label is not generated, reason: material text is not English input.")
            continue
        
        if contains_non_english_chars(reg_no):
            warnings.append(f"{identifier} label is not generated, reason: REG # is not English input.")
            continue
        
        if per_no and contains_non_english_chars(per_no):
            warnings.append(f"{identifier} label is not generated, reason: PER # is not English input.")
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


def generate_labels_to_files(
    template_path: str, 
    data_path: str, 
    output_dir: str, 
    generate_pdf: bool = True
) -> Tuple[List[str], List[str]]:
    """
    Generate label files from template and data files (file-based I/O).
    
    Args:
        template_path: Path to the SVG template file
        data_path: Path to the Excel data file
        output_dir: Directory to save generated labels
        generate_pdf: Whether to also generate PDF files
        
    Returns:
        Tuple of (svg_paths, pdf_paths)
    """
    svg_dir = os.path.join(output_dir, 'svg')
    pdf_dir = os.path.join(output_dir, 'pdf')
    os.makedirs(svg_dir, exist_ok=True)
    if generate_pdf and HAS_CAIROSVG:
        os.makedirs(pdf_dir, exist_ok=True)
    
    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()
    
    df = pd.read_excel(data_path)
    
    svg_files, pdf_files = generate_labels_from_dataframe(template_content, df, generate_pdf)
    
    svg_paths = []
    pdf_paths = []
    
    for svg_filename, svg_content in svg_files:
        svg_path = os.path.join(svg_dir, svg_filename)
        with open(svg_path, 'w', encoding='utf-8') as f:
            f.write(svg_content)
        svg_paths.append(svg_path)
        print(f"✅ Generated SVG: {svg_filename}")
    
    for pdf_filename, pdf_bytes in pdf_files:
        pdf_path = os.path.join(pdf_dir, pdf_filename)
        with open(pdf_path, 'wb') as f:
            f.write(pdf_bytes)
        pdf_paths.append(pdf_path)
        print(f"✅ Generated PDF: {pdf_filename}")
    
    return svg_paths, pdf_paths


def main():
    """CLI entry point."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    template_path = os.path.join(script_dir, 'template', 'law_label.svg')
    data_path = os.path.join(script_dir, 'data', 'law_label_data.xlsx')
    output_dir = os.path.join(script_dir, 'output')
    
    print("=" * 60)
    print("Law Label Generator (SVG + PDF)")
    print("=" * 60)
    print(f"Template: {template_path}")
    print(f"Data: {data_path}")
    print(f"Output: {output_dir}")
    print("=" * 60)
    
    if not os.path.exists(template_path):
        print(f"❌ Error: Template file not found: {template_path}")
        return
    
    if not os.path.exists(data_path):
        print(f"❌ Error: Data file not found: {data_path}")
        return
    
    svg_files, pdf_files = generate_labels_to_files(template_path, data_path, output_dir, generate_pdf=True)
    
    print("=" * 60)
    print(f"✅ Successfully generated {len(svg_files)} SVG label(s)")
    if HAS_CAIROSVG:
        print(f"✅ Successfully generated {len(pdf_files)} PDF label(s)")
    print(f"📁 SVG output: {os.path.join(output_dir, 'svg')}")
    if HAS_CAIROSVG:
        print(f"📁 PDF output: {os.path.join(output_dir, 'pdf')}")


if __name__ == "__main__":
    main()