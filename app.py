import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel & CSV Merger", page_icon="📊")

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

st.title("📊 File Merger (Excel & CSV)")
st.write("Upload multiple Excel and/or CSV files below to combine them into one master file.")

# UPGRADE 1: Added "csv" to the accepted file types
uploaded_files = st.file_uploader(
    "Upload Excel or CSV files", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}"
)

if uploaded_files:
    st.info(f"📂 {len(uploaded_files)} file(s) uploaded and ready to merge.")

    if st.button("⚙️ Merge Files"):
        dataframes_list = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, file in enumerate(uploaded_files):
            try:
                # UPGRADE 2: Check the file extension and read accordingly
                if file.name.lower().endswith('.csv'):
                    df = pd.read_csv(file)
                else:
                    df = pd.read_excel(file)
                    
                df['Source_File'] = file.name
                dataframes_list.append(df)
            except Exception as e:
                st.error(f"Could not read {file.name}: {e}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))
            status_text.text(f"Processing {file.name}...")

        if dataframes_list:
            status_text.text("Combining files...")
            master_df = pd.concat(dataframes_list, ignore_index=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                master_df.to_excel(writer, index=False)
            processed_data = output.getvalue()

            st.success(f"✅ Successfully merged {len(uploaded_files)} files! (Total Rows: {len(master_df)})")
            
            st.download_button(
                label="⬇️ Download Master Excel File",
                data=processed_data,
                file_name="Master_Combined_File.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No data could be extracted.")

    st.divider()
    
    if st.button("🗑️ Clear Files & Start Over"):
        st.session_state.uploader_key += 1
        st.rerun()
