import os
import re
import pandas as pd
import pdfplumber
import shutil

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        # Loop through all pages and concatenate their text
        for page in pdf.pages:
            text += page.extract_text()
        return text

# Function to extract generator and cost data from text
def extract_data_from_text(text, file_name, folder_path):
    # Define fallback directory for files that can't be processed
    directory = 'not extracted'
    new_folder_path = os.path.join(folder_path, directory)
    os.makedirs(new_folder_path, exist_ok=True)

    # Define section header markers
    start_markers = [
        'Appendix E. - Cost Allocation Per Request',
        'Appendix E. Cost Allocation Per Request'
    ]
    end_markers = [
        'Appendix F. - Cost Allocation Per Upgrade Facility',
        'Appendix F. Cost Allocation by Upgrade',
        'Appendix F. - Cost Allocation Per Request'
    ]

    # Initialize indices
    start_index = -1
    end_index = -1

    # Find start of the relevant section
    for marker in start_markers:
        start_index = text.find(marker)
        if start_index != -1:
            break

    # Find end of the relevant section
    for marker in end_markers:
        end_index = text.find(marker)
        if end_index != -1:
            break

    # If section headers not found, move file to fallback folder and skip
    if start_index == -1 or end_index == -1:
        print('Start or end index not found.')
        src = os.path.join(folder_path, file_name)
        dst = os.path.join(new_folder_path, file_name)
        shutil.move(src, dst)
        return None

    # Extract section between start and end markers
    extracted_text = text[start_index:end_index]

    # Pattern to capture Gen numbers (e.g., GEN-2022-123N)
    gen_pattern = r"\n((GEN|ASGI)[ -]20\d{2}-\d{3}(?:[NOS\d]+)?)\n"
    # Alternate pattern for older format (commented out)
    # gen_pattern = r"\n(G0\d{1}-\d{2,3}(?:[GR\d]+)?)\n"

    # Pattern to locate the word "Total"
    total_pattern = r"(Total)"

    # Pattern to find cost pairs like "$1,000.00 $500.00"
    cost_pattern = r"\$(\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s+\$(\d{1,3}(?:,\d{3})*(?:\.\d+)?)"

    # Find all generator occurrences
    gen_matches = list(re.finditer(gen_pattern, extracted_text))
    data = []
    search_start_index = 0  # Used to avoid reprocessing sections

    # Iterate over each generator block
    for i, match in enumerate(gen_matches):
        gen_number = match.group(1)
        gen_start_index = match.start()

        # Skip if revisiting old section
        if gen_start_index < search_start_index:
            continue

        # Find the next "Total" after this generator block
        total_match = re.search(total_pattern, extracted_text[gen_start_index:])
        if total_match:
            total_index = gen_start_index + total_match.start()
            search_start_index = total_index  # Move past this block for next iteration
        else:
            total_index = len(extracted_text)  # Use end of text if "Total" not found

        # Extract section from Gen start to "Total"
        section_text = extracted_text[gen_start_index:total_index]

        # Extract cost matches (tuples of two costs)
        cost_matches = re.findall(cost_pattern, section_text)

        # Select only the second value in each cost pair (likely upgrade costs)
        selected_costs = [cost[1] for cost in cost_matches]
        print(selected_costs)

        # Convert to float and sum up total upgrade cost
        total_upgrade_cost = sum(float(cost.replace(',', '')) for cost in selected_costs)

        # Store result in a dict
        data.append({
            'File Name': file_name,
            'Gen Number': gen_number,
            'Total Upgrade Cost': total_upgrade_cost
        })

    # Return as DataFrame if data exists, else return None
    if data:
        return pd.DataFrame(data)
    else:
        return None

# Function to apply extraction to all PDFs in a folder
def parse_pdfs_in_folder(folder_path):
    dfs = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, file_name)
            print(file_name)  # Log which file is being processed
            text = extract_text_from_pdf(pdf_path)
            df = extract_data_from_text(text, file_name, folder_path)
            if df is not None:
                dfs.append(df)

    # Combine all individual dataframes into one
    if dfs:
        return pd.concat(dfs, ignore_index=True)
    else:
        return None

# Define input and output file paths
folder_path = '../downloaded_studies/Feasibility/Group Studies'
output_file_path = '../test_files/tuc_extract_feas_t1_redo.xlsx'

# Run the parser
df = parse_pdfs_in_folder(folder_path)
print(df)

# Save final combined DataFrame to Excel
if df is not None:
    df.to_excel(output_file_path, index=False)
    print(f"Data extracted and saved to '{output_file_path}'.")
else:
    print("No data extracted.")
