import os
import fitz
import pandas as pd

# Create an empty list to store the extracted information
data = []

# Set the path of the top-level folder to search for PDF files
top_level_folder = r"D:\Coding\Python codes"

# Loop through all PDF files in the current directory and its subdirectories
for dirpath, dirnames, filenames in os.walk(top_level_folder):
    for filename in filenames:
        if filename.endswith('.pdf'):
            # Open the PDF file
            file_path = os.path.join(dirpath, filename)
            doc = fitz.open(file_path)

            # Get the first page of the PDF
            page = doc[0]

            # Extract the text from the page
            text = page.get_text()

            # Check if the first page contains the "sertifika" string
            if "sertifika" in text.lower():
                # Find the address information
                address_lines = []
                coordinates = []
                for i, line in enumerate(text.splitlines()):
                    if i==8:
                        coordinates.append(line)
                    if i==9:
                        address_lines.append(line)
                    if i==10:
                        address_lines.append(line)

                # Add the extracted information to the list
                data.append({'File Name': filename, 'Address': ', '.join(address_lines), 'Coordinates': ','.join(coordinates)})

# Create a pandas dataframe from the extracted information
df = pd.DataFrame(data)

# Save the dataframe to an Excel file
df.to_excel('addresses.xlsx', index=False)
