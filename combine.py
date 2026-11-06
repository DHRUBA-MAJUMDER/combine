import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Merger", page_icon="📊")
st.title("📊 Excel File Merger")
st.write("Upload multiple Excel files below to combine them into one master file.")

# 1. Create a file uploader widget in the web UI
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Merge Files"):
        dataframes_list = []
        
        # Progress bar for visual feedback
        progress_bar = st.progress(0)
        status_text = st.empty()

        # 2. Process uploaded files directly from memory (no folders needed)
        for i, file in enumerate(uploaded_files):
            try:
                df = pd.read_excel(file)
                df['Source_File'] = file.name
                dataframes_list.append(df)
            except Exception as e:
                st.error(f"Could not read {file.name}: {e}")
            
            # Update progress
            progress_bar.progress((i + 1) / len(uploaded_files))
            status_text.text(f"Processing {file.name}...")

        if dataframes_list:
            status_text.text("Combining files...")
            master_df = pd.concat(dataframes_list, ignore_index=True)
            
            # 3. Save the master dataframe to memory so the user can download it
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                master_df.to_excel(writer, index=False)
            processed_data = output.getvalue()

            st.success(f"Successfully merged {len(uploaded_files)} files! (Total Rows: {len(master_df)})")
            
            # 4. Create the download button
            st.download_button(
                label="⬇️ Download Master Excel File",
                data=processed_data,
                file_name="Master_Combined_File.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("No data could be extracted.")