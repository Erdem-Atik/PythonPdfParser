import os
import fitz
import pandas as pd

# Open the PDF file
doc = fitz.open('e-5.pdf')

# Get the first page of the PDF
page = doc[0]

# Extract the text from the page
text = page.get_text()

for i, line in enumerate(text.splitlines()):
    if i==9 or i==10:
        print(line)
