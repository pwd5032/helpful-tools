import json
import pandas as pd

# Define the function to extract measures
def extract_measures(bim_file_path, output_excel_path):
    # Load the BIM file content
    with open(bim_file_path, 'r', encoding='utf-8') as file:
        bim_content = json.load(file)

    # Extract measures
    measures = []
    for table in bim_content['model']['tables']:
        for measure in table.get('measures', []):
            measure_name = measure.get('name', 'N/A')
            measure_expression = measure.get('expression', 'N/A')
            measure_data_type = measure.get('dataType', 'N/A')
            measures.append({
                'Measure Name': measure_name,
                'Definition': measure_expression,
                'Data Type': measure_data_type
            })

    # Create a DataFrame
    measures_df = pd.DataFrame(measures)

    # Write to Excel
    measures_df.to_excel(output_excel_path, index=False)

# Usage
bim_file_path = 'path_to_your_model.bim'
output_excel_path = 'extracted_measures.xlsx'
extract_measures(bim_file_path, output_excel_path)
