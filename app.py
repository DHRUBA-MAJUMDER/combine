import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel & CSV Merger", page_icon="📊")

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

st.title("📊 File Merger (Excel & CSV)")
st.write("Upload multiple Excel or CSV files below to combine them into one master file.")

uploaded_files = st.file_uploader(
    "Upload Excel or CSV files", 
    type=["xlsx", "csv", "xls"], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}"
)

# UPGRADE: Added layout columns for Output and Header options
col1, col2 = st.columns(2)

with col1:
    output_format = st.radio(
        "Select Output Format:", 
        ["CSV (Fastest, prevents server crash)", "Excel .xlsx"]
    )

with col2:
    # UI UPGRADE: Checkbox to handle missing headers (Govt/School reports often lack them)
    has_header = st.checkbox("My files have a Header row", value=True)
    st.caption("⚠️ **UNCHECK THIS** for files like 'CHARAIDEO_SAPEKHATI_Report' where data starts on the very first line. This stops blank rows and prevents missing data.")

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} file(s) uploaded and ready to merge.")

    if st.button("⚙️ Merge Files"):
        dataframes_list = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        # If has_header is False, we tell Pandas not to use the first row as columns
        header_val = 0 if has_header else None

        for i, file in enumerate(uploaded_files):
            try:
                file.seek(0)
                
                if file.name.lower().endswith('.csv'):
                    try:
                        df = pd.read_csv(file, header=header_val, skipinitialspace=True, on_bad_lines='skip', encoding='utf-8')
                    except UnicodeDecodeError:
                        file.seek(0)
                        df = pd.read_csv(file, header=header_val, skipinitialspace=True, on_bad_lines='skip', encoding='cp1252')
                else:
                    df = pd.read_excel(file, header=header_val)
                
                # --- AUTO-CLEANUP ---
                # Drop rows that are completely empty
                df.dropna(how='all', inplace=True)
                
                if has_header:
                    # Remove ghost 'Unnamed' columns and strip hidden spaces from headers
                    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
                    df.columns = df.columns.astype(str).str.strip()
                
                df['Source_File'] = file.name
                dataframes_list.append(df)
            except Exception as e:
                st.error(f"Could not read {file.name}: {e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))
            status_text.text(f"Processing {file.name}...")

        if dataframes_list:
            status_text.text("Combining files...")
            master_df = pd.concat(dataframes_list, ignore_index=True)
            
            # If there was no header, Pandas uses numbers (0,1,2..). We convert them to strings to avoid Excel errors.
            if not has_header:
                master_df.columns = master_df.columns.astype(str)

            status_text.text("Preparing file for download...")
            
            if "CSV" in output_format:
                processed_data = master_df.to_csv(index=False).encode('utf-8')
                file_name = "Master_Combined_File.csv"
                mime_type = "text/csv"
            else:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    master_df.to_excel(writer, index=False)
                processed_data = output.getvalue()
                file_name = "Master_Combined_File.xlsx"
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            st.success(f"✅ Successfully merged {len(uploaded_files)} files! (Total Rows: {len(master_df)})")
            
            st.download_button(
                label=f"⬇️ Download {file_name}",
                data=processed_data,
                file_name=file_name,
                mime=mime_type
            )
        else:
            st.error("No data could be extracted.")

    st.divider()
    
    if st.button("🗑️ Clear Files & Start Over"):
        st.session_state.uploader_key += 1
        st.rerun()
