import pandas as pd
import numpy as np

# Function to convert a numeric column index to an Excel-style column letter
def excel_column_name(n):
    name = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        name = chr(65 + remainder) + name
    return name

# Paths to the spreadsheets
sheet1_path = 'sheet-origin.xlsx'
sheet2_path = 'sheet-changed.xlsx'

# Initialize DataFrames
sheet1 = None
sheet2 = None

# Try to load the spreadsheets
try:
    sheet1 = pd.read_excel(sheet1_path, header=None)
except FileNotFoundError:
    print(f"The file {sheet1_path} was not found.")

try:
    sheet2 = pd.read_excel(sheet2_path, header=None)
except FileNotFoundError:
    print(f"The file {sheet2_path} was not found.")

# Proceed only if both sheets have been successfully loaded
if sheet1 is not None and sheet2 is not None:
    # Assuming that the first row of each sheet contains the headers
    # Find common columns by header values
    common_headers = set(sheet1.iloc[0]).intersection(sheet2.iloc[0])
    common_columns = [sheet1.columns[sheet1.iloc[0].tolist().index(header)] for header in common_headers]

    # Create a list to store the differences
    differences = []

    # Check all rows and columns in both spreadsheets
    for column_index in common_columns:
        # Convert to alphabetical letter, adjusting the column index by +1
        column_letter = excel_column_name(column_index + 1)

        for row in range(max(len(sheet1), len(sheet2))):
            # Get the value from sheet 1
            val1 = sheet1.iat[row, column_index] if row < len(sheet1) else np.nan

            # Get the value from sheet 2
            val2 = sheet2.iat[row, column_index] if row < len(sheet2) else np.nan
            
            # Check if one of the values is not null and if they are different
            if pd.notna(val1) or pd.notna(val2):
                if val1 != val2:
                    differences.append({
                        'Line': row + 1,  # +1 to match Excel row numbering
                        'Column': column_letter,
                        'Value in Sheet 1': val1,
                        'Value in Sheet 2': val2
                    })

    # Convert the list of differences to a DataFrame
    differences_df = pd.DataFrame(differences)
    differences_df.sort_values(by=['Line', 'Column'], inplace=True)

    # Display the differences found
    if not differences_df.empty:
        print("\nSheet 1: " + sheet1_path)
        print("Sheet 2: " + sheet2_path)
        print("\nDifferences found:")
        print(differences_df.to_string(index=False))
    else:
        print("\nNo differences found.")
else:
    print("\nComparison could not be performed due to missing file(s).")
