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
    
    #main_file_path = st.text_input("Enter the path of the main Excel file:")
    main_file_path =r"C:\Users\esen\OneDrive\OneDrive - MSCI Office 365\Documents\exercise\count checking.xlsx"
    
    
    if main_file_path:
        # Remove the double quotes if present
        main_file_path = main_file_path.strip('"')
        
        if os.path.exists(main_file_path):
            try:
                main_df = pd.read_excel(main_file_path)
                valid_issuer_ids = main_df['DMX_ISSUER_ID'].unique()
                
                uploaded_files = st.file_uploader("Select Excel Files", type="xlsx", accept_multiple_files=True)
                
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
                    
                    count_data = filtered_data.groupby(['RUN_DATE', 'DMX_ISSUER_NAME', 'DMX_ISSUER_ID']).size().reset_index(name='count')
                    
                    count_data['order'] = count_data['DMX_ISSUER_ID'].apply(lambda x: list(valid_issuer_ids).index(x))
                    count_data = count_data.sort_values(['order', 'RUN_DATE']).drop(columns='order')
                    
                    with pd.ExcelWriter(main_file_path, engine='openpyxl', mode='a') as writer:
                        count_data.to_excel(writer, sheet_name='NewSheet', index=False)
                    
                    st.success(f"Result saved to a new sheet 'NewSheet' in: {main_file_path}")
                    st.dataframe(count_data)
                else:
                    st.warning("No files selected for validation.")
            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Invalid file path or file does not exist.")
    else:
        st.error("No main file specified.")

if __name__ == "__main__":
    main()
