import os
import fitz
import pandas as pd


data = []

folder_path = 'D:/Coding/Python codes'


# Traverse all subdirectories of the folder path
for root, dirs, files in os.walk(folder_path):
    for filename in files:
        if filename.endswith('.pdf'):
            # Open the PDF file
            doc = fitz.open(os.path.join(root, filename))

            # Get the first page of the PDF
            page = doc[0]

            # Extract the text from the page
            text = page.get_text()

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