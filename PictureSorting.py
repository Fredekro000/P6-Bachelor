import os
import pandas as pd
import glob
from PIL import Image, ImageDraw, ImageFont

# Define paths
sorted_data_path = "sorted_data.xlsx"
image_folder = r"C:\Users\Bruger\Documents\GitHub\P6WEB\P6-Bachelor\Brest\1. Måned"

# Load the sorted data
sorted_data = pd.read_excel(sorted_data_path)


# Function to extract ID from image filename
def extract_id_from_filename(filename):
    return filename.split()[0]  # Extracts the first part like 'BA1' from 'BA1 (1).JPG'


# Get all image files
image_files = glob.glob(os.path.join(image_folder, "*.JPG")) + glob.glob(os.path.join(image_folder, "*.PNG"))

# List to store matched results
matched_data = []

# Iterate over images and match with sorted data
for image_path in image_files:
    image_name = os.path.basename(image_path)  # Extract file name
    image_id = extract_id_from_filename(image_name)  # Extract ID

    # Find matching row in sorted data
    match = sorted_data[sorted_data["ID"] == image_id]

    if not match.empty:
        matched_data.append({
            "Image Name": image_name,
            "Image Path": image_path,
            "ID": image_id,
            "Position": match.iloc[0]["Position"],
            "Days": match.iloc[0]["Days"],
            "Overall Score": match.iloc[0]["Overall Score"]
        })

# Add text to images
for item in matched_data:
    img = Image.open(item["Image Path"])
    draw = ImageDraw.Draw(img)

    # Try loading a readable font
    try:
        font = ImageFont.truetype("arial.ttf", 200)  # Larger font size for visibility
    except IOError:
        font = ImageFont.load_default()

    text = f"ID: {item['ID']}\nPosition: {item['Position']}\nDays: {item['Days']}\nScore: {item['Overall Score']}"
    text_position = (50, 50)  # Offset text to avoid the edge
    text_color = (255, 255, 255)  # Red color for visibility
    stroke_fill = (10, 10, 10)  # Black stroke for better contrast
    stroke_width = 20


    draw.text(text_position, text, fill=text_color, font=font, stroke_width=stroke_width, stroke_fill=stroke_fill)

    output_path = os.path.join(image_folder, "annotated_" + item["Image Name"])
    img.save(output_path)
    print(f"Annotated image saved: {output_path}")

# Print matched data
for item in matched_data:
    print(item)
