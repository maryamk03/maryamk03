import pandas as pd
import os

# Directory containing the Excel files
directory = '../downloaded_studies/Impact/XLSX'

# Initialize an empty DataFrame for the output
output_df = pd.DataFrame(columns=['File Name', 'Gen Number', 'Interconnection Cost'])

# Loop through each file in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xlsx") or filename.endswith(".xls"):
        file_path = os.path.join(directory, filename)
        # Load the Excel file
        xl = pd.ExcelFile(file_path)
        
        # Check if the sheet exists
        if 'Assigned Upgrade Costs' in xl.sheet_names:
            # Load the specific sheet into a DataFrame
            df = xl.parse('Assigned Upgrade Costs')
            
            # Determine the correct column names
            if 'Upgrade Details' in df.columns:
                upgrade_det_col = 'Upgrade Details'
            elif 'Upgrade Details ' in df.columns:
                upgrade_det_col = 'Upgrade Details '
            else:
                # If neither column is found, skip to the next file
                continue

            if 'Allocated Cost' in df.columns:
                allocated_cost_col = 'Allocated Cost'
            elif 'Allocated Cost ' in df.columns:
                allocated_cost_col = 'Allocated Cost '
            else:
                # If neither column is found, skip to the next file
                continue

            if 'Gen Number' in df.columns:
                gen_number_col = 'Gen Number'
            elif 'Gen Number ' in df.columns:
                gen_number_col = 'Gen Number '
            else:
                # If neither column is found, skip to the next file
                continue
            
            # Filter the DataFrame to include rows where the 'Upgrade Details' column contains 'Interconnection'
            # for 2016
            # filtered_df = df[df['Upgrade Name'].str.contains('Interconnection Costs', na=False, case=False)]

            filtered_df = df[df[upgrade_det_col].str.contains('Interconnection upgrades and cost estimates|Facilitate the interconnection',
                                                                na=False, case=False)]
            
            # Extract the necessary columns
            temp_df = pd.DataFrame({
                'Gen Number': filtered_df[gen_number_col],
                'File Name': filename,
                'Interconnection Cost': filtered_df[allocated_cost_col]
            })
            # For each sheet, if there's more than one row with the same gen_number_col, sum the values in allocated_cost_col
             # Group by 'Gen Number' and sum the 'Allocated Cost'
            grouped_df = filtered_df.groupby(gen_number_col).agg({
                allocated_cost_col: 'sum'
            }).reset_index()
            
            # Add the 'File Name' column
            grouped_df['File Name'] = filename
            
            # Rename columns to match the output DataFrame
            grouped_df.rename(columns={gen_number_col: 'Gen Number', allocated_cost_col: 'Interconnection Cost'}, inplace=True)
            
            # Append to the output DataFrame
            output_df = pd.concat([output_df, grouped_df], ignore_index=True)

# Save the output DataFrame to an Excel file
output_df.to_excel('../test_files/ic_extract_xl_v2_t3.xlsx', index=False)

print(f"Data extracted from the Excel files and saved.")
