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
                
                count_data = filtered_data.groupby(['RUN_DATE', 'DMX_ISSUER_NAME', 'DMX_ISSUER_ID']).size().reset_index(name='count')
                
                count_data['order'] = count_data['DMX_ISSUER_ID'].apply(lambda x: list(valid_issuer_ids).index(x))
                count_data = count_data.sort_values(['order', 'RUN_DATE']).drop(columns='order')
                
                output_file_path = st.text_input("Enter the path to save the output Excel file:")
                
                if output_file_path:
                    # Remove any extra quotes and spaces
                    output_file_path = output_file_path.strip().strip('"')
                    st.write(f"Output file path: {output_file_path}")  # Debugging statement
                    
                    if not output_file_path.endswith('.xlsx'):
                        st.error("Error: The output file path must end with '.xlsx'")
                        return
                    
                    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                        count_data.to_excel(writer, sheet_name='NewSheet', index=False)
                        main_df.to_excel(writer, sheet_name='MainFileData', index=False)
                    
                    st.success(f"Result saved to: {output_file_path}")
                    st.dataframe(count_data)
                else:
                    st.warning("Please enter a valid output file path.")
            else:
                st.warning("No additional files selected for validation.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("No main file uploaded.")

if __name__ == "__main__":
    main()