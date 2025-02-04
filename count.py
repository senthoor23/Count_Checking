import pandas as pd
import streamlit as st
import os

def combine_excel_files(files):
    combined_data = pd.DataFrame()

    for file in files:
        df = pd.read_excel(file)
        combined_data = pd.concat([combined_data, df], ignore_index=True)
    
    return combined_data

def main():
    st.title("Excel File Combiner and Validator")
    
    main_file = st.file_uploader("Upload the main Excel file", type="xlsx")
    
    if main_file:
        try:
            main_df = pd.read_excel(main_file)
            valid_issuer_ids = main_df['DMX_ISSUER_ID'].unique()
            
            uploaded_files = st.file_uploader("Select additional Excel files", type="xlsx", accept_multiple_files=True)
            
            if uploaded_files:
                combined_data = combine_excel_files(uploaded_files)
                
                # Check if 'DMX_ISSUER_ID' and 'RUN_DATE' are in the columns
                if 'DMX_ISSUER_ID' not in combined_data.columns or 'RUN_DATE' not in combined_data.columns:
                    st.error("Error: 'DMX_ISSUER_ID' or 'RUN_DATE' column is missing from the combined data.")
                    return
                
                filtered_data = combined_data[
                    (combined_data['DMX_ISSUER_ID'].isin(valid_issuer_ids)) &
                    (combined_data['VALIDATION'] == 'Validation Needed')
                ]
                
                filtered_data = filtered_data[
                    ~filtered_data['FISCAL_YEAR'].isin([2021, 2022])
                ]
                
                count_data = filtered_data.groupby(['RUN_DATE', 'DMX_ISSUER_NAME', 'DMX_ISSUER_ID']).size().reset_index(name='Correct_Count')
                
                # Add the 'TOTAL' column from the main file
                count_data = count_data.merge(main_df[['DMX_ISSUER_ID', 'TOTAL']], on='DMX_ISSUER_ID', how='left')
                
                count_data['order'] = count_data['DMX_ISSUER_ID'].apply(lambda x: list(valid_issuer_ids).index(x))
                count_data = count_data.sort_values(['order', 'RUN_DATE']).drop(columns='order')
                
                # Provide the output as a downloadable CSV file
                csv = count_data.to_csv(index=False)
                st.download_button(
                    label="Download output as CSV",
                    data=csv,
                    file_name='output.csv',
                    mime='text/csv',
                )
                
                st.success("Output generated successfully!")
                st.dataframe(count_data)
            else:
                st.warning("No additional files selected for validation.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("No main file uploaded.")

if __name__ == "__main__":
    main()
