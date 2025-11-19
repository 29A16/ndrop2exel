#!/usr/bin/env python3
"""
Streamlit application to convert XPS files to PDF and extract tables to Excel format.
Handles Swedish number format (comma as decimal separator).
"""

import streamlit as st
import os
import tempfile
import pandas as pd
import tabula.io as tabula
import re
import subprocess
import shutil
import zipfile
import io
from datetime import datetime

def fix_swedish_numbers(df):
    """
    Fix Swedish number format where comma is decimal separator.
    Converts strings like "4,141" to proper float 4.141
    """
    for col in df.columns:
        if df[col].dtype == 'object':  # Only process string columns
            # Try to convert Swedish format numbers
            df[col] = df[col].apply(lambda x: convert_swedish_number(x) if isinstance(x, str) else x)
    return df

def convert_swedish_number(value):
    """
    Convert Swedish number format to standard format.
    """
    if not isinstance(value, str):
        return value
    
    # Check if it looks like a number with comma
    # Pattern: optional digits, comma, digits
    if re.match(r'^\d+[,\.]\d+', value):
        # Replace comma with dot, and remove any dots used as thousands separator
        # First, identify if there are multiple separators
        comma_count = value.count(',')
        dot_count = value.count('.')
        
        if comma_count == 1 and dot_count == 0:
            # Simple case: just replace comma with dot
            try:
                return float(value.replace(',', '.'))
            except ValueError:
                return value
        elif comma_count == 1 and dot_count == 1:
            # Complex case like "173,0.71" which should be "173.071"
            # Remove the first digit after comma, replace comma with dot
            # This handles the OCR error
            value = value.replace(',', '.')
            # If we now have something like "173.0.71", remove middle digit
            parts = value.split('.')
            if len(parts) == 3 and len(parts[1]) == 1:
                value = parts[0] + '.' + parts[2]
            try:
                return float(value)
            except ValueError:
                return value
    
    return value

def merge_fragmented_tables(tables):
    """
    Merge tables that are fragments of the same table split across pages.
    """
    if len(tables) <= 1:
        return tables
    
    merged = []
    current_table = None
    
    for table in tables:
        # Skip completely empty tables
        if table.empty:
            continue
            
        # Check if this looks like a continuation (starts with Unnamed: 0 or has same columns)
        is_continuation = False
        
        if current_table is not None:
            # Check if columns match (indicating same table)
            if list(table.columns) == list(current_table.columns):
                is_continuation = True
            # Check if first column is "Unnamed: 0" (common fragment indicator)
            elif table.columns[0].startswith('Unnamed'):
                # Try to merge as continuation
                is_continuation = True
        
        if is_continuation and current_table is not None:
            # Merge with current table
            current_table = pd.concat([current_table, table], ignore_index=True)
        else:
            # Start new table
            if current_table is not None:
                merged.append(current_table)
            current_table = table.copy()
    
    # Don't forget the last table
    if current_table is not None:
        merged.append(current_table)
    
    return merged

def extract_summary_data(table, source_file):
    """
    Extract Sample, ng/ul, and 260/280 columns from a table.
    Handles various column name variations.
    Adds source file name to each row.
    """
    # Possible column name variations
    sample_cols = ['Sample', 'sample', 'Sample Name', 'sample name']
    ngul_cols = ['ng/ul', 'ng/uL', 'ng/µl', 'Concentration', 'concentration']
    ratio_cols = ['260/280', '260 / 280', 'A260/A280', 'Ratio']
    
    extracted = pd.DataFrame()
    
    # Find Sample column
    sample_col = None
    for col in table.columns:
        if any(s in str(col) for s in sample_cols):
            sample_col = col
            break
    
    # Find ng/ul column
    ngul_col = None
    for col in table.columns:
        if any(n in str(col) for n in ngul_cols):
            ngul_col = col
            break
    
    # Find 260/280 column
    ratio_col = None
    for col in table.columns:
        if any(r in str(col) for r in ratio_cols):
            ratio_col = col
            break
    
    # Extract the columns we found
    if sample_col:
        extracted['Sample'] = table[sample_col]
    if ngul_col:
        extracted['ng/ul'] = table[ngul_col]
    if ratio_col:
        extracted['260/280'] = table[ratio_col]
    
    # Add source file name as first column
    if not extracted.empty:
        extracted.insert(0, 'Source File', source_file)
        # Apply Swedish number format fix to extracted data
        extracted = fix_swedish_numbers(extracted)
    
    return extracted if not extracted.empty else None

def convert_pdf_to_excel(pdf_path, pdf_data=None):
    """Convert a single PDF to Excel with Swedish number format handling."""
    try:
        # If PDF data is provided (from uploaded file), write it to temporary file
        if pdf_data:
            with open(pdf_path, 'wb') as f:
                f.write(pdf_data)
        
        # Extract tables from PDF with better parameters
        tables = tabula.read_pdf(
            pdf_path, 
            pages='all', 
            multiple_tables=True,
            lattice=True,  # Use lattice mode for tables with lines
            guess=True     # Fall back to guess mode if lattice fails
        )
        
        if not tables:
            return None, None, f"No tables found in {os.path.basename(pdf_path)}"
        
        # Merge fragmented tables
        tables = merge_fragmented_tables(tables)
        
        # Output filename
        excel_filename = os.path.basename(pdf_path).replace('.pdf', '.xlsx')
        excel_path = os.path.join(os.path.dirname(pdf_path), excel_filename)
        
        # Collect summary data from all tables
        summary_data = []
        
        # If multiple tables, save to different sheets
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for idx, table in enumerate(tables):
                # Fix Swedish number format
                table = fix_swedish_numbers(table)
                
                # Extract summary data (pass the PDF filename without extension)
                source_name = os.path.splitext(os.path.basename(pdf_path))[0]
                summary = extract_summary_data(table, source_name)
                if summary is not None:
                    summary_data.append(summary)
                
                # Write to Excel
                sheet_name = f'Table_{idx+1}' if len(tables) > 1 else 'Sheet1'
                table.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Combine summary data if any was extracted
        combined_summary = None
        if summary_data:
            combined_summary = pd.concat(summary_data, ignore_index=True)
        
        return excel_path, combined_summary, f"Successfully processed {len(tables)} table(s)"
        
    except Exception as e:
        return None, None, f"Error processing {os.path.basename(pdf_path)}: {str(e)}"

def convert_xps_to_pdf(xps_path, xps_data):
    """
    Convert XPS file to PDF using xpstopdf (from libgxps-utils).
    Returns the PDF path and data if successful, None otherwise.
    """
    try:
        # Write XPS data to temporary file
        with open(xps_path, 'wb') as f:
            f.write(xps_data)
        
        pdf_path = xps_path.replace('.xps', '.pdf').replace('.XPS', '.pdf')
        
        # Use xpstopdf to convert (requires libgxps-utils package)
        result = subprocess.run(
            ['xpstopdf', xps_path, pdf_path],
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0 and os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                pdf_data = f.read()
            return pdf_path, pdf_data, "XPS converted to PDF successfully"
        else:
            error_msg = "Failed to convert XPS to PDF"
            if result.stderr:
                error_msg += f": {result.stderr}"
            return None, None, error_msg
            
    except FileNotFoundError:
        return None, None, "xpstopdf command not found. Please install libgxps-utils"
    except Exception as e:
        return None, None, f"Error converting XPS: {str(e)}"

def create_download_zip(files_data):
    """Create a ZIP file containing all the Excel files"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, data in files_data.items():
            zip_file.writestr(filename, data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

def main():
    st.set_page_config(
        page_title="XPS to Excel Converter",
        page_icon=None,
        layout="wide"
    )
    
    st.title("XPS to Excel Converter")
    st.markdown("Convert XPS files to PDF and extract tables to Excel format with Swedish/German number format support")
    
    # Sidebar for options
    with st.sidebar:
        st.header("Options")
        include_summary = st.checkbox("Generate Combined Summary", value=True, 
                                    help="Create a summary Excel file with Sample, ng/ul, and 260/280 data from all files")
        
        st.header("Requirements")
        st.info("""
        **For XPS files:** Requires `libgxps-utils` package
        ```bash
        sudo apt-get install libgxps-utils
        ```
        """)
    
    # File upload section
    st.header("Upload Files")
    uploaded_files = st.file_uploader(
        "Choose PDF or XPS files",
        type=['pdf', 'xps'],
        accept_multiple_files=True,
        help="Upload one or more PDF or XPS files to extract tables"
    )
    
    if uploaded_files:
        st.success(f"Uploaded {len(uploaded_files)} file(s)")
        
        # Display uploaded files
        with st.expander("View uploaded files"):
            for file in uploaded_files:
                st.write(f"• {file.name} ({file.size} bytes)")
        
        # Process files button
        if st.button("Process Files", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            results = []
            all_summaries = []
            excel_files_data = {}
            
            with tempfile.TemporaryDirectory() as temp_dir:
                for idx, uploaded_file in enumerate(uploaded_files):
                    progress = (idx + 1) / len(uploaded_files)
                    progress_bar.progress(progress)
                    status_text.text(f"Processing {uploaded_file.name}...")
                    
                    file_extension = uploaded_file.name.lower().split('.')[-1]
                    temp_file_path = os.path.join(temp_dir, uploaded_file.name)
                    
                    try:
                        if file_extension == 'xps':
                            # Convert XPS to PDF first
                            pdf_path, pdf_data, message = convert_xps_to_pdf(temp_file_path, uploaded_file.getvalue())
                            
                            if pdf_path and pdf_data:
                                # Process the converted PDF
                                excel_path, summary, excel_message = convert_pdf_to_excel(pdf_path, pdf_data)
                                message = f"XPS→PDF→Excel: {excel_message}"
                            else:
                                excel_path, summary, excel_message = None, None, message
                                
                        elif file_extension == 'pdf':
                            # Process PDF directly
                            excel_path, summary, message = convert_pdf_to_excel(temp_file_path, uploaded_file.getvalue())
                        
                        # Store results
                        if excel_path and os.path.exists(excel_path):
                            with open(excel_path, 'rb') as f:
                                excel_data = f.read()
                            excel_filename = os.path.basename(excel_path)
                            excel_files_data[excel_filename] = excel_data
                            
                            if summary is not None:
                                all_summaries.append(summary)
                            
                            results.append({
                                'file': uploaded_file.name,
                                'status': 'Success',
                                'message': message,
                                'excel_file': excel_filename
                            })
                        else:
                            results.append({
                                'file': uploaded_file.name,
                                'status': 'Failed',
                                'message': message,
                                'excel_file': None
                            })
                    
                    except Exception as e:
                        results.append({
                            'file': uploaded_file.name,
                            'status': 'Error',
                            'message': str(e),
                            'excel_file': None
                        })
                
                # Create combined summary if requested and data is available
                if include_summary and all_summaries:
                    try:
                        combined_summary = pd.concat(all_summaries, ignore_index=True)
                        summary_filename = f"combined_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        summary_path = os.path.join(temp_dir, summary_filename)
                        
                        with pd.ExcelWriter(summary_path, engine='openpyxl') as writer:
                            combined_summary.to_excel(writer, index=False, sheet_name='Summary')
                        
                        with open(summary_path, 'rb') as f:
                            excel_files_data[summary_filename] = f.read()
                        
                        st.success(f"Created combined summary with {len(combined_summary)} rows")
                    except Exception as e:
                        st.error(f"Error creating combined summary: {str(e)}")
            
            progress_bar.progress(1.0)
            status_text.text("Processing complete!")
            
            # Display results
            st.header("Processing Results")
            results_df = pd.DataFrame(results)
            st.dataframe(results_df, use_container_width=True)
            
            # Download section
            if excel_files_data:
                st.header("Download Results")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("Individual Files")
                    for filename, data in excel_files_data.items():
                        st.download_button(
                            label=f"{filename}",
                            data=data,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                with col2:
                    st.subheader("All Files (ZIP)")
                    if len(excel_files_data) > 1:
                        zip_data = create_download_zip(excel_files_data)
                        st.download_button(
                            label="Download All as ZIP",
                            data=zip_data,
                            file_name=f"excel_extractions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                            mime="application/zip"
                        )
                    else:
                        st.info("Only one file available - use individual download above")
                
                # Success summary
                success_count = len([r for r in results if r['status'] == 'Success'])
                st.success(f"Successfully processed {success_count}/{len(uploaded_files)} files")
                
            else:
                st.error("No Excel files were generated. Please check the error messages above.")
    
    else:
        # Instructions when no files are uploaded
        st.info("Upload PDF or XPS files to get started")
        
        with st.expander("How it works"):
            st.markdown("""
            **This application:**
            1. **XPS Files**: Converts XPS to PDF using `xpstopdf` (requires libgxps-utils)
            2. **PDF Files**: Extracts tables using advanced table detection
            3. **Number Format**: Handles Swedish format (comma as decimal separator)
            4. **Table Merging**: Automatically merges fragmented tables from multiple pages
            5. **Excel Export**: Saves each file's tables to separate Excel sheets
            6. **Summary**: Optionally creates a combined summary with Sample, ng/ul, and 260/280 data
            
            **Supported formats:**
            - PDF files with tables
            - XPS files (converted to PDF first)
            
            **Output:**
            - Individual Excel files for each input file
            - Combined summary Excel file (optional)
            - ZIP download for multiple files
            """)

if __name__ == '__main__':
    main()