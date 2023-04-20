import os
import fitz
import pandas as pd
import re

# Create an empty list to store the extracted information
data = []

# Set the path of the top-level folder to search for PDF files
top_level_folder = r"D:\CODING\PYTHON\BTK_BAZ_İSTASYON\Guv_Sert_Parser\PDFs"

# Loop through all PDF files in the current directory and its subdirectories
for dirpath, dirnames, filenames in os.walk(top_level_folder):
    
    for filename in filenames:

        if filename.endswith('.pdf') or filename.endswith('.PDF'):
           
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
                capitalized_sentences = []
                coordinates = []
                for i, line in enumerate(text.splitlines()):
 
                    if '°'in line:      
                        coordinates.append(line)
                    if line.isupper() and not('°'in line):
                        capitalized_sentences.append(line)

                sentence_to_remove =["BİLGİ TEKNOLOJİLERİ VE İLETİŞİM KURUMU","T.C.", "GÜVENLİK SERTİFİKASI", "VODAFONE TELEKOMÜNİKASYON ANONİM ŞİRKETİ","HÜCRESEL SİSTEM 2N",": VODAFONE TELEKOMÜNİKASYON A. Ş.",": HÜCRESEL SİSTEM 4.5N"]
                filtered_sentences = [s for s in capitalized_sentences if not any(p in s for p in sentence_to_remove)]
                pattern_BTK = r"BTK \d+"
                
                address_lines = [s for s in filtered_sentences if not re.search(pattern_BTK, s)]

                # Add the extracted information to the list
                data.append({'File Name': filename[0:8], 'Address': ', '.join(address_lines), 'Coordinates': ','.join(coordinates)})


# Create a pandas dataframe from the extracted information
df = pd.DataFrame(data)

# Save the dataframe to an Excel file
df.to_excel('SiteID_AdresKoord.xlsx', index=False)

