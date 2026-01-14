#!/usr/bin/env python3
"""
Label Generator - Streamlit Web App

A web interface for generating labels from Excel data.
Upload an xlsx file and download generated PDF labels as a zip file.

Generates both Label 2 and Label 19 from the same data.
"""

import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import datetime

# Import label generation functions
from generate_label2 import (
    generate_label2_from_dataframe as generate_label2,
    HAS_CAIROSVG
)
from generate_label19 import (
    generate_label19_from_dataframe as generate_label19
)

# Page configuration
st.set_page_config(
    page_title="Label Generator",
    page_icon="🏷️",
    layout="centered"
)

# Title
st.title("🏷️ Label Generator")

# Description
st.markdown("""
**Upload your Excel file to generate both Label 2 and Label 19 at once.**

**Expected Excel format:**
- Column 1: Product code (used for filename)
- Column 2: Material composition text (max 15 lines)
- Column 3: REG. No
- Column 4: PER. No (optional)
- Column 5: Firm
- Column 6: Origin (CN/VN)
""")

# Check if cairosvg is available
if not HAS_CAIROSVG:
    st.error("❌ cairosvg is not installed. PDF generation is not available.")
    st.stop()

# Get template paths
script_dir = os.path.dirname(os.path.abspath(__file__))
template2_path = os.path.join(script_dir, 'template', 'label2.svg')
template19_path = os.path.join(script_dir, 'template', 'label19.svg')

# Check if templates exist
templates_missing = []
if not os.path.exists(template2_path):
    templates_missing.append("label2.svg")
if not os.path.exists(template19_path):
    templates_missing.append("label19.svg")

if templates_missing:
    st.error(f"❌ Template file(s) not found: {', '.join(templates_missing)}")
    st.stop()

# Load templates
with open(template2_path, 'r', encoding='utf-8') as f:
    template2_content = f.read()
with open(template19_path, 'r', encoding='utf-8') as f:
    template19_content = f.read()

# File uploader
st.subheader("📁 Upload Data File")
uploaded_file = st.file_uploader(
    "Select an Excel file (.xlsx)", 
    type=["xlsx"],
    help="Upload the Excel file containing label data"
)

if uploaded_file is not None:
    # Read and preview data
    try:
        df = pd.read_excel(uploaded_file)
        
        st.success(f"✅ File loaded: {uploaded_file.name} ({len(df)} rows)")
        
        # Generate button
        if st.button("🚀 Generate Labels", type="primary", use_container_width=True):
            with st.spinner("Generating labels..."):
                all_pdf_files = []
                all_warnings = []
                
                # Generate Label 2
                pdf_files_2, warnings_2 = generate_label2(
                    template2_content, 
                    df
                )
                all_pdf_files.extend(pdf_files_2)
                all_warnings.extend(warnings_2)
                
                # Generate Label 19
                pdf_files_19, warnings_19 = generate_label19(
                    template19_content,
                    df
                )
                all_pdf_files.extend(pdf_files_19)
                all_warnings.extend(warnings_19)
                
                # Display warnings if any
                if all_warnings:
                    st.subheader("⚠️ Warnings")
                    for warning in all_warnings:
                        st.warning(warning)
                
                if not all_pdf_files:
                    st.error("❌ No labels were generated. Check your data file.")
                else:
                    # Create zip file in memory
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        # Add PDF files
                        for filename, content in all_pdf_files:
                            zf.writestr(filename, content)
                    
                    zip_buffer.seek(0)
                    
                    # Count labels by type
                    label2_count = len(pdf_files_2)
                    label19_count = len(pdf_files_19)
                    
                    # Show success message
                    st.success(f"✅ Generated {label2_count} Label 2 PDFs and {label19_count} Label 19 PDFs!")
                    
                    # Download button
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    zip_filename = f"labels_{timestamp}.zip"
                    
                    st.download_button(
                        label="📥 Download All Labels (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=zip_filename,
                        mime="application/zip",
                        use_container_width=True
                    )
    
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")

# Footer
st.divider()
st.caption("Label Generator v1.2 | Upload Excel → Generate Label 2 & Label 19 → Download ZIP")
