import pandas as pd
from fillpdf import fillpdfs

# Load the Excel file
df = pd.read_excel('001Test.xlsx', header=None)  # Replace with your Excel file path

# Define the column ranges for each group
ranges = {
    'Delta': range(3, 6),       # Columns D (3rd index) to F (5th index)
    'Theta': range(6, 10),      # Columns G (6th index) to J (9th index)
    'Alpha': range(10, 15),     # Columns K (10th index) to O (14th index)
    'SMR': range(14, 18),       # Columns O (14th index) to R (17th index)
    'LowBeta': range(18, 23),   # Columns S (18th index) to W (22nd index)
    'HighBeta': range(23, 33)   # Columns X (23rd index) to AG (32nd index)
}

# Function to calculate averages for a given row range
def calculate_averages(row_range_left, row_range_right):
    averages = {}
    for group, col_range in ranges.items():
        left_values = []
        right_values = []
        
        for col_idx in col_range:
            column_data = df.iloc[:, col_idx]
            
            # Select specific rows for Left and Right averages
            left_values.extend(column_data.iloc[row_range_left])
            right_values.extend(column_data.iloc[row_range_right])
        
        # Calculate averages
        left_average = round(pd.Series(left_values).mean(), 1)
        right_average = round(pd.Series(right_values).mean(), 1)

        # Store averages in variables
        averages[f'{group}_left'] = left_average
        averages[f'{group}_right'] = right_average

    return averages

# Calculate OE1 averages
OE1_left_rows = range(11, 70, 2)  # Rows 12, 14, ..., 70
OE1_right_rows = range(12, 71, 2)  # Rows 13, 15, ..., 71
OE1_averages = calculate_averages(OE1_left_rows, OE1_right_rows)

# Calculate OE2 averages
OE2_left_rows = range(91, 150, 2)  # Rows 92, 94, ..., 150
OE2_right_rows = range(92, 151, 2)  # Rows 93, 95, ..., 151
OE2_averages = calculate_averages(OE2_left_rows, OE2_right_rows)

# Calculate CE averages
CE_left_rows = range(171, 230, 2)  # Rows 172, 174, ..., 230
CE_right_rows = range(172, 231, 2)  # Rows 173, 175, ..., 231
CE_averages = calculate_averages(CE_left_rows, CE_right_rows)

# Calculate combined averages (Combine)
combined_left_rows = list(OE1_left_rows) + list(OE2_left_rows) + list(CE_left_rows)
combined_right_rows = list(OE1_right_rows) + list(OE2_right_rows) + list(CE_right_rows)
Combine_averages = calculate_averages(combined_left_rows, combined_right_rows)

# Calculate additional fields based on the provided logic
additional_fields = {
    'Delta/SMR_left': round(Combine_averages['Delta_left'] / Combine_averages['SMR_left'], 2),
    'Delta/SMR_right': round(Combine_averages['Delta_right'] / Combine_averages['SMR_right'], 2),
    'Delta/LowBeta_left': round(Combine_averages['Delta_left'] / Combine_averages['LowBeta_left'], 2),
    'Delta/LowBeta_right': round(Combine_averages['Delta_right'] / Combine_averages['LowBeta_right'], 2),
    'Delta/HighBeta_left': round(Combine_averages['Delta_left'] / Combine_averages['HighBeta_left'], 2),
    'Delta/HighBeta_right': round(Combine_averages['Delta_right'] / Combine_averages['HighBeta_right'], 2),
    'Delta/Alpha_left': round(Combine_averages['Delta_left'] / Combine_averages['Alpha_left'], 2),
    'Delta/Alpha_right': round(Combine_averages['Delta_right'] / Combine_averages['Alpha_right'], 2),
    'Alpha/HighBeta_left': round(Combine_averages['Alpha_left'] / Combine_averages['HighBeta_left'], 2),
    'Alpha/HighBeta_right': round(Combine_averages['Alpha_right'] / Combine_averages['HighBeta_right'], 2),
    'SMR/HighBeta_left': round(Combine_averages['SMR_left'] / Combine_averages['HighBeta_left'], 2),
    'SMR/HighBeta_right': round(Combine_averages['SMR_right'] / Combine_averages['HighBeta_right'], 2),
    'Alpha/LowBeta_left': round(Combine_averages['Alpha_left'] / Combine_averages['LowBeta_left'], 2),
    'Alpha/LowBeta_right': round(Combine_averages['Alpha_right'] / Combine_averages['LowBeta_right'], 2),
    'Theta/LowBeta_left': round(Combine_averages['Theta_left'] / Combine_averages['LowBeta_left'], 2),
    'Theta/LowBeta_right': round(Combine_averages['Theta_right'] / Combine_averages['LowBeta_right'], 2),
    'LowBeta_left': str(Combine_averages['LowBeta_left']),
    'LowBeta_right': str(Combine_averages['LowBeta_right']),
    'HighBeta_left': str(Combine_averages['HighBeta_left']),
    'HighBeta_right': str(Combine_averages['HighBeta_right']),
    'Alpha_left': str(Combine_averages['Alpha_left']),
    'Alpha_right': str(Combine_averages['Alpha_right']),
    'LowBeta_left1': str(Combine_averages['LowBeta_left']),  # Same as LowBeta_left
    'LowBeta_right1': str(Combine_averages['LowBeta_right']), 
}

# Map all averages and additional fields to PDF field names
field_values = {}

# Add OE1 values
field_values.update({f"OE1_{key}": str(value) for key, value in OE1_averages.items()})

# Add OE2 values
field_values.update({f"OE2_{key}": str(value) for key, value in OE2_averages.items()})

# Add CE values
field_values.update({f"CE_{key}": str(value) for key, value in CE_averages.items()})

# Add combined (Combine) values
field_values.update({f"CO_{key}": str(value) for key, value in Combine_averages.items()})

# Add additional fields
field_values.update({key: str(value) for key, value in additional_fields.items()})

# Input and output file paths for the PDF
input_pdf = "input/bfm-template.pdf"  # Replace with your input PDF path
output_pdf = "output/bfm-template_fill.pdf"  # Final non-editable (flattened) PDF

# Fill and flatten the PDF
temp_filled_pdf = "output/temp_filled.pdf"
fillpdfs.write_fillable_pdf(input_pdf, temp_filled_pdf, field_values)
fillpdfs.flatten_pdf(temp_filled_pdf, output_pdf)

# Final confirmation
print(f"OE1, OE2, CE, Combine averages, and additional fields filled and non-editable PDF saved as: {output_pdf}")
