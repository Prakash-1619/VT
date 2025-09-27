import pandas as pd
import os
import requests
from urllib.parse import unquote, urlparse
from datetime import datetime
import re
import streamlit as st
import zipfile
from io import BytesIO

# Set page config
st.set_page_config(
    page_title="File Renaming Tool",
    page_icon="📁",
    layout="wide"
)

def main():
    st.title("📁 Advanced File Renaming Tool")
    st.markdown("Upload Excel, filter data, and download renamed files as ZIP")
    
    # Initialize session state
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'filtered_df' not in st.session_state:
        st.session_state.filtered_df = None
    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = []
    
    # File upload section
    st.header("1. Upload Excel File")
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file", 
        type=['xlsx', 'xls'],
        help="Upload your Excel file with file links"
    )
    
    if uploaded_file is not None:
        try:
            if st.session_state.df is None:
                df = pd.read_excel(uploaded_file)
                st.session_state.df = df
                st.session_state.filtered_df = df.copy()
                st.success(f"✅ File uploaded successfully! ({len(df)} records, {len(df.columns)} columns)")
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
    
    if st.session_state.df is not None:
        df = st.session_state.df
        
        # Display data overview
        col1, col2, col3 = st.columns([2, 1, 1])
        
        with col1:
            st.subheader("Data Preview")
            st.dataframe(st.session_state.filtered_df.head(8), use_container_width=True)
        
        with col2:
            st.subheader("Dataset Info")
            st.metric("Total Records", len(df))
            st.metric("Filtered Records", len(st.session_state.filtered_df))
            st.metric("Columns", len(df.columns))
        
        with col3:
            st.subheader("Actions")
            if st.button("🔄 Reset Filters", use_container_width=True):
                st.session_state.filtered_df = df.copy()
                st.rerun()
        
        # Filtering section
        st.header("2. Filter Data")
        with st.expander("🔍 Filter Records", expanded=True):
            # Create filters for each column
            filters = {}
            filter_cols = st.columns(3)
            
            for i, column in enumerate(df.columns):
                col_idx = i % 3
                with filter_cols[col_idx]:
                    if df[column].dtype == 'object':
                        # Text filters
                        unique_vals = df[column].dropna().unique()
                        if len(unique_vals) <= 20:
                            selected_vals = st.multiselect(
                                f"{column}",
                                options=unique_vals,
                                default=[],
                                key=f"filter_{column}"
                            )
                            if selected_vals:
                                filters[column] = selected_vals
                        else:
                            search_term = st.text_input(
                                f"Search {column}",
                                key=f"search_{column}"
                            )
                            if search_term:
                                filters[column] = search_term
                    else:
                        # Numeric/date filters
                        if pd.api.types.is_numeric_dtype(df[column]):
                            min_val = float(df[column].min()) if not pd.isna(df[column].min()) else 0
                            max_val = float(df[column].max()) if not pd.isna(df[column].max()) else 100
                            val_range = st.slider(
                                f"{column} Range",
                                min_val, max_val, (min_val, max_val),
                                key=f"range_{column}"
                            )
                            filters[column] = val_range
        
        # Apply filters button
        if st.button("Apply Filters", type="primary"):
            filtered_df = df.copy()
            
            for column, filter_value in filters.items():
                if isinstance(filter_value, list):  # Multi-select
                    if filter_value:
                        filtered_df = filtered_df[filtered_df[column].isin(filter_value)]
                elif isinstance(filter_value, tuple):  # Range
                    filtered_df = filtered_df[
                        (filtered_df[column] >= filter_value[0]) & 
                        (filtered_df[column] <= filter_value[1])
                    ]
                elif isinstance(filter_value, str):  # Text search
                    if filter_value:
                        filtered_df = filtered_df[
                            filtered_df[column].astype(str).str.contains(filter_value, case=False, na=False)
                        ]
            
            st.session_state.filtered_df = filtered_df
            st.rerun()
        
        # Column selection section
        st.header("3. Select Columns for Filename")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Required Columns")
            date_column = st.selectbox(
                "📅 Date Column*",
                options=df.columns,
                index=0,
                help="Select column containing dates"
            )
            
            serial_source = st.radio(
                "Serial Number Source",
                ["Auto-generate", "Select from column"],
                help="Auto-generate or select from existing column"
            )
            
            if serial_source == "Select from column":
                serial_column = st.selectbox(
                    "🔢 Serial Number Column",
                    options=df.columns,
                    index=1 if len(df.columns) > 1 else 0
                )
            else:
                serial_column = None
            
            bill_type_column = st.selectbox(
                "💰 Bill Type Column*",
                options=df.columns,
                index=2 if len(df.columns) > 2 else 0,
                help="Select column containing bill types"
            )
        
        with col2:
            st.subheader("Additional Columns")
            member_column = st.selectbox(
                "👤 Member Column*",
                options=df.columns,
                index=3 if len(df.columns) > 3 else 0,
                help="Select column containing member names"
            )
            
            purpose_column = st.selectbox(
                "🎯 Purpose Column*",
                options=df.columns,
                index=4 if len(df.columns) > 4 else 0,
                help="Select column containing purposes"
            )
            
            vendor_column = st.selectbox(
                "🏢 Vendor Column*",
                options=df.columns,
                index=5 if len(df.columns) > 5 else 0,
                help="Select column containing vendor names"
            )
            
            # Optional event column
            event_option = st.checkbox("Include Event Column in filename")
            if event_option:
                event_column = st.selectbox(
                    "🎪 Event Column",
                    options=df.columns,
                    index=6 if len(df.columns) > 6 else 0,
                    help="Optional: Include event information in filename"
                )
            else:
                event_column = None
        
        # File link column
        st.subheader("File Link Column")
        link_column = st.selectbox(
            "📎 File Link/URL Column*",
            options=df.columns,
            index=0,
            help="Select column containing file URLs"
        )
        
        # Storage settings
        st.header("4. Storage Settings")
        storage_col1, storage_col2 = st.columns([2, 1])
        
        with storage_col1:
            folder_name = st.text_input(
                "📁 Folder Name",
                value="renamed_files",
                help="Name for the output folder"
            )
        
        with storage_col2:
            st.write("")
            st.write("")
            if st.button("Create Folder", type="secondary"):
                os.makedirs(folder_name, exist_ok=True)
                st.success(f"Folder ready: {folder_name}")
        
        # Preview section
        st.header("5. Preview")
        if st.button("👀 Preview Filenames", type="secondary"):
            preview_data = preview_filenames(
                st.session_state.filtered_df,
                date_column,
                serial_column,
                bill_type_column,
                member_column,
                purpose_column,
                vendor_column,
                event_column
            )
            st.dataframe(preview_data, use_container_width=True)
        
        # Process files section
        st.header("6. Process & Download")
        
        process_col1, process_col2 = st.columns([3, 1])
        
        with process_col1:
            st.write(f"**Files to process:** {len(st.session_state.filtered_df)}")
            st.write(f"**Output folder:** {folder_name}")
        
        with process_col2:
            if st.button("🚀 PROCESS FILES", type="primary", use_container_width=True):
                result = process_files(
                    st.session_state.filtered_df,
                    folder_name,
                    link_column,
                    date_column,
                    serial_column,
                    bill_type_column,
                    member_column,
                    purpose_column,
                    vendor_column,
                    event_column
                )
                
                if result['successful'] > 0:
                    # Create ZIP file for download
                    zip_buffer = create_zip_folder(folder_name)
                    
                    st.success(f"✅ Processed {result['successful']} files successfully!")
                    
                    # Download button
                    st.download_button(
                        label="📥 Download All Files as ZIP",
                        data=zip_buffer.getvalue(),
                        file_name=f"{folder_name}.zip",
                        mime="application/zip"
                    )
                
                # Show results
                show_processing_results(result)

def preview_filenames(df, date_col, serial_col, bill_type_col, member_col, 
                     purpose_col, vendor_col, event_col):
    """Preview how files will be renamed"""
    preview_data = []
    date_serial_map = {}
    
    for index, row in df.iterrows():
        try:
            # Skip if essential data is missing
            if pd.isna(row[date_col]) or pd.isna(row[bill_type_col]) or pd.isna(row[member_col]):
                preview_data.append({
                    'Row': index + 2,
                    'Status': '❌ Missing required data',
                    'Filename': 'Cannot generate'
                })
                continue
            
            # Process date
            date_obj = pd.to_datetime(row[date_col], errors='coerce')
            if pd.isna(date_obj):
                preview_data.append({
                    'Row': index + 2,
                    'Status': '❌ Invalid date',
                    'Filename': 'Cannot generate'
                })
                continue
            
            date_str = date_obj.strftime('%Y%m%d')
            
            # Serial number
            if serial_col and not pd.isna(row[serial_col]):
                serial_no = str(int(row[serial_col])).zfill(3)
            else:
                # Auto-generate
                if date_str not in date_serial_map:
                    date_serial_map[date_str] = 1
                else:
                    date_serial_map[date_str] += 1
                serial_no = str(date_serial_map[date_str]).zfill(3)
            
            # Get values
            bill_type = sanitize_text(row[bill_type_col])
            member = sanitize_text(row[member_col])
            purpose = sanitize_text(row[purpose_col]) if purpose_col and not pd.isna(row[purpose_col]) else "Purpose"
            vendor = sanitize_text(row[vendor_col]) if vendor_col and not pd.isna(row[vendor_col]) else "Vendor"
            
            # Build filename
            filename_parts = [date_str, serial_no, bill_type, member, purpose, vendor]
            
            # Add event if specified
            if event_col and not pd.isna(row[event_col]):
                event = sanitize_text(row[event_col])
                filename_parts.append(event)
            
            filename = "-".join(filename_parts)
            
            preview_data.append({
                'Row': index + 2,
                'Status': '✅ Ready',
                'Filename': filename + "[.extension]"
            })
            
        except Exception as e:
            preview_data.append({
                'Row': index + 2,
                'Status': f'❌ Error: {str(e)}',
                'Filename': 'Cannot generate'
            })
    
    return pd.DataFrame(preview_data)

def process_files(df, folder_path, link_col, date_col, serial_col, bill_type_col, 
                 member_col, purpose_col, vendor_col, event_col):
    """Process and download files"""
    
    # Create folder
    os.makedirs(folder_path, exist_ok=True)
    
    # Initialize
    results = []
    date_serial_map = {}
    processed_files = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for index, (_, row) in enumerate(df.iterrows()):
        try:
            # Update progress
            progress = (index + 1) / len(df)
            progress_bar.progress(progress)
            status_text.text(f"Processing {index + 1}/{len(df)}...")
            
            # Skip if missing required data
            if (pd.isna(row.get(link_col)) or pd.isna(row.get(date_col)) or 
                pd.isna(row.get(bill_type_col)) or pd.isna(row.get(member_col))):
                results.append({
                    'row': index + 2,
                    'status': 'Failed - Missing data',
                    'filename': 'Skipped'
                })
                continue
            
            # Process date
            date_obj = pd.to_datetime(row[date_col], errors='coerce')
            if pd.isna(date_obj):
                results.append({
                    'row': index + 2,
                    'status': 'Failed - Invalid date',
                    'filename': 'Skipped'
                })
                continue
            
            date_str = date_obj.strftime('%Y%m%d')
            
            # Serial number
            if serial_col and not pd.isna(row[serial_col]):
                try:
                    serial_no = str(int(float(row[serial_col]))).zfill(3)
                except:
                    serial_no = "001"
            else:
                # Auto-generate
                if date_str not in date_serial_map:
                    date_serial_map[date_str] = 1
                else:
                    date_serial_map[date_str] += 1
                serial_no = str(date_serial_map[date_str]).zfill(3)
            
            # Get values
            bill_type = sanitize_text(row[bill_type_col])
            member = sanitize_text(row[member_col])
            purpose = sanitize_text(row[purpose_col]) if purpose_col and not pd.isna(row[purpose_col]) else "Purpose"
            vendor = sanitize_text(row[vendor_col]) if vendor_col and not pd.isna(row[vendor_col]) else "Vendor"
            
            # Build filename
            filename_parts = [date_str, serial_no, bill_type, member, purpose, vendor]
            
            # Add event if specified
            if event_col and not pd.isna(row[event_col]):
                event = sanitize_text(row[event_col])
                filename_parts.append(event)
            
            base_filename = "-".join(filename_parts)
            
            # Download and save file
            url = row[link_col]
            try:
                response = requests.get(url, stream=True, timeout=60)
                response.raise_for_status()
                
                # Determine extension
                extension = get_file_extension(response, url)
                final_filename = base_filename + extension
                final_path = os.path.join(folder_path, final_filename)
                
                # Save file
                with open(final_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                
                results.append({
                    'row': index + 2,
                    'status': 'Success',
                    'filename': final_filename
                })
                processed_files.append(final_path)
                
            except Exception as e:
                results.append({
                    'row': index + 2,
                    'status': f'Failed - Download error',
                    'filename': base_filename
                })
                
        except Exception as e:
            results.append({
                'row': index + 2,
                'status': f'Failed - {str(e)}',
                'filename': 'Error'
            })
    
    progress_bar.empty()
    status_text.empty()
    
    st.session_state.processed_files = processed_files
    
    return {
        'total': len(df),
        'successful': len([r for r in results if r['status'] == 'Success']),
        'failed': len([r for r in results if r['status'] != 'Success']),
        'results': results,
        'folder_path': folder_path
    }

def sanitize_text(text):
    """Sanitize text for filename"""
    if pd.isna(text):
        return "Unknown"
    text = str(text)
    # Remove invalid characters and limit length
    sanitized = re.sub(r'[<>:"/\\|?*]', '', text).strip()
    return sanitized.replace(' ', '_')[:30]

def get_file_extension(response, url):
    """Get correct file extension"""
    content_type = response.headers.get('content-type', '').lower()
    
    if 'pdf' in content_type:
        return '.pdf'
    elif 'jpeg' in content_type or 'jpg' in content_type:
        return '.jpg'
    elif 'png' in content_type:
        return '.png'
    elif 'gif' in content_type:
        return '.gif'
    else:
        # Try from URL
        parsed_url = urlparse(url)
        url_ext = os.path.splitext(parsed_url.path)[1].lower()
        if url_ext in ['.jpg', '.jpeg', '.png', '.pdf', '.gif']:
            return url_ext
        return '.jpg'  # Default to jpg for images

def create_zip_folder(folder_path):
    """Create ZIP file of the folder"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zip_file.write(file_path, os.path.relpath(file_path, folder_path))
    zip_buffer.seek(0)
    return zip_buffer

def show_processing_results(result):
    """Display processing results"""
    st.header("📊 Processing Results")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Files", result['total'])
    with col2:
        st.metric("Successful", result['successful'])
    with col3:
        st.metric("Failed", result['failed'])
    
    # Show results table
    results_df = pd.DataFrame(result['results'])
    st.dataframe(results_df, use_container_width=True)
    
    st.info(f"📁 Files saved to: `{result['folder_path']}`")

if __name__ == "__main__":
    main()
