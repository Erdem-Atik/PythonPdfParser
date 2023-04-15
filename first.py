import os
import fitz
import pandas as pd

# Create an empty list to store the extracted information
data = []

# Loop through all PDF files in the current directory
for filename in os.listdir('.'):
    if filename.endswith('.pdf'):
        # Open the PDF file
        doc = fitz.open(filename)

        # Get the first page of the PDF
        page = doc[0]

        # Extract the text from the page
        text = page.get_text()

        # Find the address information
        address_lines = []
        for i, line in enumerate(text.splitlines()):
            if i==9:
                address_lines.append(line)
            if i==10:
                address_lines.append(line)

        # Add the extracted information to the list
        data.append({'File Name': filename, 'Address': ', '.join(address_lines)})

# Create a pandas dataframe from the extracted information
df = pd.DataFrame(data)

# Save the dataframe to an Excel file
df.to_excel('addresses.xlsx', index=False)
