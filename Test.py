import os
import pandas as pd

# Load Excel file
file_path = "CEPE 2012 inspection AF Efficacy Atlantic.xlsx"
xls = pd.ExcelFile(file_path)


# Function to clean the ID (remove hyphens, numbers, and lowercase)
def clean_string(s):
    return ''.join([char for char in s if char.isalpha()]).lower()


# Picture coding matrix (mapping from ID to position)
picture_coding = [
    ["BB2", "BN3", "BP1", "BD1"],
    ["BA3", "BC2", "BA1", "BE1"],
    ["BC3", "BE2", "BN2", "BB1"],
    ["BD2", "BP2", "BB3", "BC1"],
    ["BA2", "BE3", "BD3", "BN1"]
]

# Create position map for ID matching
position_map = {id_: f"{chr(65 + r)}{c + 1}" for r, row in enumerate(picture_coding) for c, id_ in enumerate(row)}

# Observation days and starting row offsets
observation_days = [36, 62, 90, 122, 155, 184]
start_row = 9 - 2  # First item starts here (Row 9)
item_spacing = 12  # 8 lines + 4 padding

# List for extracted data
extracted_data = []

# Iterate through all sheets and extract data
for sheet in ["Line1", "Line2", "Line3", "Line4", "Line5"]:
    df = xls.parse(sheet)

    for i in range(4):  # 4 items per sheet
        row = start_row + (item_spacing * i)
        ID = str(df.iat[row, 0]).strip()  # ID from column A
        Position = position_map.get(ID, "Unknown")  # Find position from picture coding matrix
        Days = observation_days[0]
        Overall_score = df.iloc[row + 7, 2]  # Get overall score

        extracted_data.append({
            "ID": ID,
            "Position": Position,
            "Days": Days,
            "Overall Score": Overall_score
        })

# Convert extracted data into a DataFrame
df_extracted = pd.DataFrame(extracted_data)

# Save the sorted data to a new Excel file
output_file = "sorted_data.xlsx"
df_extracted.to_excel(output_file, index=False)

print(f"Extracted data has been saved to {output_file}.")
