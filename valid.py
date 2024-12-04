import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import os

# Function to perform fuzzy matching
def is_match(school_name, check_schools, threshold):
    matches = process.extract(school_name, check_schools, scorer=fuzz.ratio)
    return any(match[1] >= threshold for match in matches)

# Streamlit UI for file upload and threshold input
st.title('Maxcare India School Names Matching Tool')

# Upload input Excel sheet and check sheet
input_file = st.file_uploader("Upload Input Sheet (Excel)", type=['xlsx'])
check_file = st.file_uploader("Upload Check Sheet (Excel)", type=['xlsx'])

# If the files are uploaded
if input_file is not None and check_file is not None:
    # Read the input and check sheets
    input_df = pd.read_excel(input_file)
    check_df = pd.read_excel(check_file)

    # Ensure the column names are correctly recognized (case sensitive check)
    if 'Name' not in input_df.columns or 'INSTITUTE' not in check_df.columns:
        st.error("The input sheet must have a 'Name' column and the check sheet must have an 'INSTITUTE' column.")
    else:
        # Slider to set the threshold for fuzzy matching
        threshold = st.slider("Select Fuzzy Matching Threshold (%)", 0, 100, 85)

        # Extract school names from both sheets
        input_schools = input_df['Name'].tolist()
        check_schools = check_df['INSTITUTE'].tolist()

        # Find unmatched schools
        unmatched_rows = input_df[~input_df['Name'].apply(lambda x: is_match(x, check_schools, threshold))]

        # Save the output file with the same name but with '_processed' suffix
        output_filename = os.path.splitext(input_file.name)[0] + '_processed.xlsx'
        unmatched_rows.to_excel(output_filename, index=False)

        # Provide a download link for the processed file
        st.write(f"Processed file saved as: {output_filename}")
        st.download_button(
            label="Download Processed File",
            data=open(output_filename, 'rb').read(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

