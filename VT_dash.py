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
                event_column,
                link_column
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
                with st.spinner("Processing files..."):
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
                        mime="application/zip",
                        key="download_zip"
                    )
                
                # Show results
                show_processing_results(result)

def preview_filenames(df, date_col, serial_col, bill_type_col, member_col, 
                     purpose_col, vendor_col, event_col, link_col):
    """Preview how files will be renamed"""
    preview_data = []
    date_serial_map = {}
    
    for index, row in df.iterrows():
        try:
            # Skip if essential data is missing
            if (pd.isna(row.get(date_col)) or pd.isna(row.get(bill_type_col)) or 
                pd.isna(row.get(member_col)) or pd.isna(row.get(link_col))):
                preview_data.append({
                    'Row': index + 2,
                    'Status': '❌ Missing required data',
                    'Filename': 'Cannot generate',
                    'File Type': 'Unknown'
                })
                continue
            
            # Process date
            date_obj = pd.to_datetime(row[date_col], errors='coerce')
            if pd.isna(date_obj):
                preview_data.append({
                    'Row': index + 2,
                    'Status': '❌ Invalid date',
                    'Filename': 'Cannot generate',
                    'File Type': 'Unknown'
                })
                continue
            
            date_str = date_obj.strftime('%Y%m%d')
            
            # Serial number
            if serial_col and not pd.isna(row.get(serial_col)):
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
            purpose = sanitize_text(row[purpose_col]) if purpose_col and not pd.isna(row.get(purpose_col)) else "Purpose"
            vendor = sanitize_text(row[vendor_col]) if vendor_col and not pd.isna(row.get(vendor_col)) else "Vendor"
            
            # Build filename
            filename_parts = [date_str, serial_no, bill_type, member, purpose, vendor]
            
            # Add event if specified
            if event_col and not pd.isna(row.get(event_col)):
                event = sanitize_text(row[event_col])
                filename_parts.append(event)
            
            filename = "-".join(filename_parts)
            
            # Try to detect file type from URL
            url = row[link_col]
            file_type = detect_file_type_from_url(url)
            
            preview_data.append({
                'Row': index + 2,
                'Status': '✅ Ready',
                'Filename': filename + file_type,
                'File Type': file_type.replace('.', '').upper()
            })
            
        except Exception as e:
            preview_data.append({
                'Row': index + 2,
                'Status': f'❌ Error: {str(e)}',
                'Filename': 'Cannot generate',
                'File Type': 'Unknown'
            })
    
    return pd.DataFrame(preview_data)

def process_files(df, folder_path, link_col, date_col, serial_col, bill_type_col, 
                 member_col, purpose_col, vendor_col, event_col):
    """Process and download files with proper file type detection"""
    
    # Create folder
    os.makedirs(folder_path, exist_ok=True)
    
    # Clear existing files in the folder
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            os.remove(file_path)
    
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
                    'status': 'Failed - Missing required data',
                    'filename': 'Skipped',
                    'file_size': '0 KB'
                })
                continue
            
            # Process date
            date_obj = pd.to_datetime(row[date_col], errors='coerce')
            if pd.isna(date_obj):
                results.append({
                    'row': index + 2,
                    'status': 'Failed - Invalid date',
                    'filename': 'Skipped',
                    'file_size': '0 KB'
                })
                continue
            
            date_str = date_obj.strftime('%Y%m%d')
            
            # Serial number
            if serial_col and not pd.isna(row.get(serial_col)):
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
            purpose = sanitize_text(row[purpose_col]) if purpose_col and not pd.isna(row.get(purpose_col)) else "Purpose"
            vendor = sanitize_text(row[vendor_col]) if vendor_col and not pd.isna(row.get(vendor_col)) else "Vendor"
            
            # Build filename
            filename_parts = [date_str, serial_no, bill_type, member, purpose, vendor]
            
            # Add event if specified
            if event_col and not pd.isna(row.get(event_col)):
                event = sanitize_text(row[event_col])
                filename_parts.append(event)
            
            base_filename = "-".join(filename_parts)
            
            # Download and save file
            url = str(row[link_col]).strip()
            try:
                # First, try to get headers to determine file type
                head_response = requests.head(url, timeout=10, allow_redirects=True)
                content_type = head_response.headers.get('content-type', '').lower()
                
                # Determine extension from content type
                extension = get_file_extension_from_content_type(content_type, url)
                
                final_filename = base_filename + extension
                final_path = os.path.join(folder_path, final_filename)
                
                # Download the actual file
                response = requests.get(url, stream=True, timeout=60, allow_redirects=True)
                response.raise_for_status()
                
                # Check if we got actual content
                content_length = response.headers.get('content-length')
                if content_length and int(content_length) == 0:
                    raise Exception("Empty file content")
                
                # Save file
                with open(final_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:  # filter out keep-alive chunks
                            f.write(chunk)
                
                # Verify file was actually written and has content
                file_size = os.path.getsize(final_path)
                if file_size == 0:
                    os.remove(final_path)
                    raise Exception("Downloaded file is empty")
                
                file_size_kb = round(file_size / 1024, 1)
                
                results.append({
                    'row': index + 2,
                    'status': 'Success',
                    'filename': final_filename,
                    'file_size': f'{file_size_kb} KB',
                    'file_type': extension.upper().replace('.', '')
                })
                processed_files.append(final_path)
                
            except Exception as e:
                results.append({
                    'row': index + 2,
                    'status': f'Failed - {str(e)}',
                    'filename': base_filename,
                    'file_size': '0 KB',
                    'file_type': 'Unknown'
                })
                
        except Exception as e:
            results.append({
                'row': index + 2,
                'status': f'Failed - {str(e)}',
                'filename': 'Error',
                'file_size': '0 KB',
                'file_type': 'Unknown'
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

def detect_file_type_from_url(url):
    """Detect file type from URL for preview"""
    try:
        url = str(url).lower()
        if '.pdf' in url:
            return '.pdf'
        elif '.jpg' in url or '.jpeg' in url:
            return '.jpg'
        elif '.png' in url:
            return '.png'
        elif '.gif' in url:
            return '.gif'
        else:
            return '.file'
    except:
        return '.file'

def get_file_extension_from_content_type(content_type, url):
    """Get correct file extension from content type and URL"""
    content_type = str(content_type).lower()
    url = str(url).lower()
    
    # Priority to content type detection
    if 'pdf' in content_type:
        return '.pdf'
    elif 'jpeg' in content_type or 'jpg' in content_type:
        return '.jpg'
    elif 'png' in content_type:
        return '.png'
    elif 'gif' in content_type:
        return '.gif'
    elif 'image' in content_type:
        return '.jpg'  #default for images
    
    # Fallback to URL detection
    if '.pdf' in url:
        return '.pdf'
    elif '.jpg' in url or '.jpeg' in url:
        return '.jpg'
    elif '.png' in url:
        return '.png'
    elif '.gif' in url:
        return '.gif'
    
    # Final fallback
    return '.file'

def create_zip_folder(folder_path):
    """Create ZIP file of the folder"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                # Only add non-empty files
                if os.path.getsize(file_path) > 0:
                    zip_file.write(file_path, os.path.basename(file_path))
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
    
    # Show results table with file types and sizes
    results_df = pd.DataFrame(result['results'])
    
    # Add file type summary
    if not results_df.empty:
        st.subheader("File Type Summary")
        file_types = results_df[results_df['status'] == 'Success']['file_type'].value_counts()
        for file_type, count in file_types.items():
            st.write(f"- {file_type}: {count} files")
    
    st.dataframe(results_df, use_container_width=True)
    
    st.info(f"📁 Files saved to: `{result['folder_path']}`")
    
    if result['failed'] > 0:
        st.error("Some files failed to download. Check the URLs and try again.")

if __name__ == "__main__":
    main()
