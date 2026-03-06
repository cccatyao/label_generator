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
from generate_label4 import (
    generate_label4_from_dataframe as generate_label4
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
**Upload your Excel file to generate Label 2, Label 19, and Label 4 at once.**

**Expected Excel format:**
- Column 1: Product code (used for filename)
- Column 2: Material composition text (max 15 lines)
- Column 3: REG. No
- Column 4: PER. No (optional)
- Column 5: Firm
- Column 6: Origin (CN/VN/KHM)
- Column 7: Washing Material (for Label 4)
- Column 8: Washing Guide (for Label 4 icons/text)
""")

# Check if cairosvg is available
if not HAS_CAIROSVG:
    st.error("❌ cairosvg is not installed. PDF generation is not available.")
    st.stop()

# Get template paths
script_dir = os.path.dirname(os.path.abspath(__file__))
template2_path = os.path.join(script_dir, 'template', 'label2.svg')
template19_path = os.path.join(script_dir, 'template', 'label19.svg')
template4_path = os.path.join(script_dir, 'template', 'label4.svg')

# Check if templates exist
templates_missing = []
if not os.path.exists(template2_path):
    templates_missing.append("label2.svg")
if not os.path.exists(template19_path):
    templates_missing.append("label19.svg")
if not os.path.exists(template4_path):
    templates_missing.append("label4.svg")

if templates_missing:
    st.error(f"❌ Template file(s) not found: {', '.join(templates_missing)}")
    st.stop()

# Load templates
with open(template2_path, 'r', encoding='utf-8') as f:
    template2_content = f.read()
with open(template19_path, 'r', encoding='utf-8') as f:
    template19_content = f.read()
with open(template4_path, 'r', encoding='utf-8') as f:
    template4_content = f.read()

# Initialize session state for persisting results
if 'warnings' not in st.session_state:
    st.session_state.warnings = []
if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None
if 'zip_filename' not in st.session_state:
    st.session_state.zip_filename = None
if 'label2_count' not in st.session_state:
    st.session_state.label2_count = 0
if 'label19_count' not in st.session_state:
    st.session_state.label19_count = 0
if 'label4_count' not in st.session_state:
    st.session_state.label4_count = 0
if 'last_uploaded_file' not in st.session_state:
    st.session_state.last_uploaded_file = None

# File uploader
st.subheader("📁 Upload Data File")
uploaded_file = st.file_uploader(
    "Select an Excel file (.xlsx)", 
    type=["xlsx"],
    help="Upload the Excel file containing label data"
)

if uploaded_file is not None:
    # Check if a new file was uploaded - clear previous results
    if st.session_state.last_uploaded_file != uploaded_file.name:
        st.session_state.last_uploaded_file = uploaded_file.name
        st.session_state.warnings = []
        st.session_state.zip_data = None
        st.session_state.zip_filename = None
        st.session_state.label2_count = 0
        st.session_state.label19_count = 0
        st.session_state.label4_count = 0
    
    # Read and preview data
    try:
        df = pd.read_excel(uploaded_file)
        
        st.success(f"✅ File loaded: {uploaded_file.name} ({len(df)} rows)")
        
        # Generate button
        if st.button("🚀 Generate Labels", type="primary", use_container_width=True):
            with st.spinner("Generating labels..."):
                # Import validation function from centralized validation module
                from validation import validate_record_for_labels
                
                all_pdf_files = []
                all_warnings = []
                pdf_files_2 = []
                pdf_files_19 = []
                pdf_files_4 = []
                
                # Get column names
                columns = df.columns.tolist()
                materials_col = columns[1]
                reg_no_col = columns[2]
                per_no_col = columns[3] if len(columns) > 3 else None
                
                # First pass: validate all records
                valid_indices = []
                for index, row in df.iterrows():
                    materials_text = str(row[materials_col]) if pd.notna(row[materials_col]) else ""
                    reg_no = str(row[reg_no_col]) if pd.notna(row[reg_no_col]) else ""
                    identifier = str(row[columns[0]]) if pd.notna(row[columns[0]]) else f"label_{index}"
                    
                    # Get PER. No if column exists
                    per_no = ""
                    if per_no_col and per_no_col in row:
                        per_no = str(row[per_no_col]) if pd.notna(row[per_no_col]) else ""
                    
                    # Validate record for both labels
                    is_valid, validation_errors = validate_record_for_labels(
                        materials_text, reg_no, per_no, identifier
                    )
                    
                    if not is_valid:
                        # Add all validation errors as warnings
                        for error in validation_errors:
                            all_warnings.append(error)
                    else:
                        valid_indices.append(index)
                
                # Create a filtered dataframe with only valid records
                if valid_indices:
                    df_valid = df.iloc[valid_indices].reset_index(drop=True)
                    
                    # Generate Label 2 for valid records only
                    pdf_files_2, warnings_2 = generate_label2(
                        template2_content, 
                        df_valid
                    )
                    all_pdf_files.extend(pdf_files_2)
                    all_warnings.extend(warnings_2)
                    
                    # Generate Label 19 for valid records only
                    pdf_files_19, warnings_19 = generate_label19(
                        template19_content,
                        df_valid
                    )
                    all_pdf_files.extend(pdf_files_19)
                    all_warnings.extend(warnings_19)

                # Generate Label 4 independently from label2/label19 row gating
                pdf_files_4, warnings_4 = generate_label4(
                    template4_content,
                    df
                )
                all_pdf_files.extend(pdf_files_4)
                all_warnings.extend(warnings_4)
                
                # Store warnings in session state
                st.session_state.warnings = all_warnings
                
                if not all_pdf_files:
                    st.session_state.zip_data = None
                    st.session_state.zip_filename = None
                    st.session_state.label2_count = 0
                    st.session_state.label19_count = 0
                    st.session_state.label4_count = 0
                else:
                    # Create zip file in memory
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        # Add PDF files
                        for filename, content in all_pdf_files:
                            zf.writestr(filename, content)
                    
                    zip_buffer.seek(0)
                    
                    # Store results in session state
                    st.session_state.zip_data = zip_buffer.getvalue()
                    st.session_state.label2_count = len(pdf_files_2)
                    st.session_state.label19_count = len(pdf_files_19)
                    st.session_state.label4_count = len(pdf_files_4)
                    
                    # Generate filename
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    st.session_state.zip_filename = f"labels_{timestamp}.zip"
        
        # Display warnings if any exist in session state (persists after download)
        if st.session_state.warnings:
            st.subheader("⚠️ Warnings")
            for warning in st.session_state.warnings:
                st.warning(warning)
        
        # Display download button if zip data exists (persists after download)
        if st.session_state.zip_data is not None:
            # Show success message
            st.success(
                "✅ Generated "
                f"{st.session_state.label2_count} Label 2 PDFs, "
                f"{st.session_state.label19_count} Label 19 PDFs, and "
                f"{st.session_state.label4_count} Label 4 PDFs!"
            )
            
            # Download button
            st.download_button(
                label="📥 Download All Labels (ZIP)",
                data=st.session_state.zip_data,
                file_name=st.session_state.zip_filename,
                mime="application/zip",
                use_container_width=True
            )
        elif st.session_state.warnings and not st.session_state.zip_data:
            # Only show error if we have warnings but no data (generation was attempted)
            if st.session_state.last_uploaded_file == uploaded_file.name and len(st.session_state.warnings) > 0:
                st.error("❌ No labels were generated. Check your data file.")
    
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")

# Footer
st.divider()
st.caption("Label Generator v1.3 | Upload Excel → Generate Label 2/19/4 → Download ZIP")
