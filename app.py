import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel & CSV Merger", page_icon="📊")

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

st.title("📊 File Merger (Excel & CSV)")
st.write("Upload multiple Excel and/or CSV files below to combine them into one master file.")

uploaded_files = st.file_uploader(
    "Upload Excel or CSV files", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}"
)

output_format = st.radio(
    "Select Output Format:", 
    ["CSV (Recommended for 10+ files - Very Fast)", "Excel .xlsx (May crash server if files are too large)"]
)

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} file(s) uploaded and ready to merge.")

    if st.button("⚙️ Merge Files"):
        dataframes_list = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, file in enumerate(uploaded_files):
            try:
                # 1. Read the file (skipping bad lines if the CSV is malformed)
                if file.name.lower().endswith('.csv'):
                    df = pd.read_csv(file, skipinitialspace=True, on_bad_lines='skip')
                else:
                    df = pd.read_excel(file)
                
                # --- AUTO-CLEANUP START ---
                # A. Drop rows that are completely blank
                df.dropna(how='all', inplace=True)
                
                # B. Drop random "Unnamed" ghost columns that systems often export
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                
                # C. Strip invisible spaces from column headers so they match perfectly
                df.columns = df.columns.astype(str).str.strip()
                # --- AUTO-CLEANUP END ---
                
                df['Source_File'] = file.name
                dataframes_list.append(df)
            except Exception as e:
                st.error(f"Could not read {file.name}: {e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))
            status_text.text(f"Processing {file.name}...")

        if dataframes_list:
            status_text.text("Combining files...")
            master_df = pd.concat(dataframes_list, ignore_index=True)
            
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
