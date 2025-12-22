#!/usr/bin/env python3
"""
Label Generator - Streamlit Web App

A web interface for generating labels from Excel data.
Upload an xlsx file and download generated PDF labels as a zip file.
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

# Page configuration
st.set_page_config(
    page_title="Label Generator",
    page_icon="🏷️",
    layout="centered"
)

# Title
st.title("🏷️ Label Generator")

# Label type selector
st.subheader("📋 Select Label Type")
label_type = st.selectbox(
    "Choose which label to generate:",
    options=["Label 2", "Label 19"],
    index=0,
    help="Select the type of label you want to generate"
)

# Label type configurations
LABEL_CONFIGS = {
    "Label 2": {
        "template": "label2.svg",
        "generator": generate_label2,
        "description": """
**Expected Excel format for Label 2:**
- Column 1: Product code (used for filename)
- Column 2: Material composition text (max 15 lines)
- Column 3: REG. No
- Column 4: PER. No (optional)
- Column 5: Firm
- Column 6: Origin (CN/VN)
""",
        "zip_prefix": "label2"
    },
    "Label 19": {
        "template": None,  # TODO: Add template path
        "generator": None,  # TODO: Import from generate_label19
        "description": """
**Label 19:** Coming soon - template and generator not yet configured.
""",
        "zip_prefix": "label19"
    }
}

# Get current label config
config = LABEL_CONFIGS[label_type]

# Show description for selected label type
st.markdown(config["description"])

# Check if cairosvg is available
if not HAS_CAIROSVG:
    st.error("❌ cairosvg is not installed. PDF generation is not available.")
    st.stop()

# Check if label type is implemented
if config["generator"] is None:
    st.warning("⚠️ This label type is not yet implemented.")
    st.stop()

# Get template path
script_dir = os.path.dirname(os.path.abspath(__file__))
template_path = os.path.join(script_dir, 'template', config["template"])

# Check if template exists
if not os.path.exists(template_path):
    st.error(f"❌ Template file not found: {template_path}")
    st.stop()

# Load template
with open(template_path, 'r', encoding='utf-8') as f:
    template_content = f.read()

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
                # Generate labels in memory (PDF only)
                pdf_files, warnings = config["generator"](
                    template_content, 
                    df
                )
                
                # Display warnings if any
                if warnings:
                    st.subheader("⚠️ Warnings")
                    for warning in warnings:
                        st.warning(warning)
                
                if not pdf_files:
                    st.error("❌ No labels were generated. Check your data file.")
                else:
                    # Create zip file in memory
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        # Add PDF files
                        for filename, content in pdf_files:
                            zf.writestr(filename, content)
                    
                    zip_buffer.seek(0)
                    
                    # Show success message
                    st.success(f"✅ Generated {len(pdf_files)} PDF labels!")
                    
                    # Download button
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    zip_filename = f"{config['zip_prefix']}_{timestamp}.zip"
                    
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
st.caption("Label Generator v1.1 | Select Label Type → Upload Excel → Generate Labels → Download ZIP")

