import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Merger", page_icon="📊")

# This creates a dynamic key to easily reset the file uploader
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

st.title("📊 Excel File Merger")
st.write("Upload multiple Excel files below to combine them into one master file.")

# The uploader now uses the dynamic key from session state
uploaded_files = st.file_uploader(
    "Upload Excel files", 
    type=["xlsx"], 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}"
)

if uploaded_files:
    # UI UPGRADE 1: Show exactly how many files were uploaded
    st.info(f"📂 {len(uploaded_files)} file(s) uploaded and ready to merge.")

    if st.button("⚙️ Merge Files"):
        dataframes_list = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, file in enumerate(uploaded_files):
            try:
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
    
    # UI UPGRADE 2: Clear button to reset the uploader
    if st.button("🗑️ Clear Files & Start Over"):
        # Changing the session state key forces the uploader widget to reset
        st.session_state.uploader_key += 1
        st.rerun()
