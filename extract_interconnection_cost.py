import os
import re
import pandas as pd
import pdfplumber
import shutil

# Function to extract all text from a given PDF file
def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        # Loop through all pages and extract text
        for page in pdf.pages:
            text += page.extract_text()
        return text

# Function to clean the extracted text
def clean_text(text):
    # Replace newline characters with spaces
    text = re.sub(r'\n', ' ', text)
    # Collapse multiple spaces into one
    text = re.sub(r'\s+', ' ', text)
    # Merge hyphenated words split across lines
    text = re.sub(r'(\w+)-\s+(\w+)', r'\1\2', text)
    return text.strip()

# Function to extract structured data from cleaned text
def extract_data_from_text(text, file_name, folder_path):
    # Set up fallback directory for PDFs where data can't be extracted
    directory = 'not extracted'
    new_folder_path = os.path.join(folder_path, directory)
    os.makedirs(new_folder_path, exist_ok=True)

    # List of potential headers marking the start of the relevant section
    start_markers = [
        'Appendix E. - Cost Allocation Per Request',
        'E. Cost Allocation Per Request',
        'E: Cost Allocation per Interconnection Request',
        'Appendix E. Cost Allocation Per Request'
    ]

    # List of headers marking the end of the relevant section
    end_markers = [
        'Appendix F. - Cost Allocation Per Upgrade Facility',
        'Appendix F. Cost Allocation by Upgrade',
        'Appendix F. - Cost Allocation Per Request'
    ]

    # Initialize indices
    start_index = -1
    end_index = -1

    # Find the first occurrence of any valid start marker
    for marker in start_markers:
        start_index = text.find(marker)
        if start_index != -1:
            break

    # Find the first occurrence of any valid end marker
    for marker in end_markers:
        end_index = text.find(marker)
        if end_index != -1:
            break
    
    # If either marker wasn't found, move the file to 'not extracted' and return None
    if start_index == -1 or end_index == -1:
        print('Start or end index not found.')
        src = os.path.join(folder_path, file_name)
        dst = os.path.join(new_folder_path, file_name)
        shutil.move(src, dst)
        return None

    # Slice the relevant section of text between the found markers
    extracted_text = text[start_index:end_index]

    # Clean the sliced text
    cleaned_text = clean_text(extracted_text)
    print(cleaned_text)

    # Regex pattern to extract Generator ID and Interconnection cost
    pattern = r'((GEN|ASGI)\d{2}-\d{2,3}(?:[NOS]\d{0,3})?).*?Interconnection?\**?.*?\$(\d+(?:,\d+)*(?:\.\d+)?)'
    
    # Find all matches using the pattern
    matches = re.findall(pattern, cleaned_text, re.DOTALL)
    print(matches)

    # Extract generator numbers
    requests = [match[0] for match in matches]
    print(requests)

    # Extract corresponding costs and convert to float
    interconnection_costs = [float(match[2].replace(',', '')) for match in matches]
    print(interconnection_costs)

    # Build a dataframe from the extracted data
    data = {'File Name': [file_name] * len(requests),
            'Gen Number': requests, 
            'Interconnection Cost': interconnection_costs}
    df = pd.DataFrame(data)

    return df

# Function to process all PDFs in a folder
def parse_pdfs_in_folder(folder_path):
    dfs = []
    # Loop through all files in the folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, file_name)
            print(file_name)
            # Extract text from the PDF
            text = extract_text_from_pdf(pdf_path)
            # Try to extract data from the text
            df = extract_data_from_text(text, file_name, folder_path)
            # If data was successfully extracted, add it to the list
            if df is not None:
                dfs.append(df)

    # Combine all extracted dataframes into one
    if dfs:
        return pd.concat(dfs, ignore_index=True)
    else:
        return None

# Set input and output paths
folder_path = '../downloaded_studies/Feasibility/Group Studies'
output_file_path = '../test_files/ic_extract.xlsx'

# Parse all PDFs in the folder
df = parse_pdfs_in_folder(folder_path)

# Save the final combined dataframe to an Excel file
if df is not None:
    df.to_excel(output_file_path, index=False)
    print(f"Data extracted and saved to '{output_file_path}'.")
else:
    print("No data extracted.")
