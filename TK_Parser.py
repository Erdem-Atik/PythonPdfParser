import os
import fitz
import pandas as pd
import re
import camelot
import tabula
from tabulate import tabulate
from PyPDF2 import PdfReader

# Set up a DataFrame to store the data
df = pd.DataFrame(columns=["File Name", "Address", "Coordinates", "Site No","Site ID","Site Name","Length","Height"])

# Set the path of the top-level folder to search for PDF files
top_level_folder = r"C:\Users\T460s\Desktop\testPDf\PDFs" # path of PDFs to parse

def parse_adres(text):
    keywords = ["mahallesi", "mah.", "sokak", "sok","KARABÜK","CADDESİ","CAD.","cad.","ada","parsel","Karabük", "Ankara", "Çorum", "Eskişehir", "Ordu", "Samsun", "Sinop", "Düzce", "Tokat", "Erzincan", "Zonguldak", "Çankırı", "Giresun", "Kastamonu", "Ağrı", "Sivas", "Şanlıurfa", "Iğdır","Safranbolu", "Çankaya", "Odunpazarı", "Mamak", "Sivrihisar", "Tepebaşı", "Altındağ", "Akkuş", "Altınordu", "Asarcık", "Ayancık", "Boyabat", "Çilimli", "Erbaa", "Fatsa", "İlkadım", "Kabadüz", "Korgan", "Saraydüzü", "Turhal", "Ünye", "Keçiören", "Ayaş", "Aybastı", "Çamaş", "Çarşamba", "Gerze", "Gölbaşı", "Gülyalı", "Eldivan", "Espiye", "Tosya", "Haymana", "Patnos", "Pursaklar", "Tekkeköy", "Yenimahalle", "Zara", "Birecik", "Sincan", "Tirebolu", "Hilvan", "Aralık", "Akıncılar", "Yıldızeli", "Etimesgut", "Karaköprü", "Siverek", "Akçakale", "Eyyübiye", "Akçakoca"]  # İçermesini istediğiniz kelimeler

    adresler = []
    for line in text.splitlines():
        for keyword in keywords:
            if keyword.lower() in line.lower():
                adresler.append(line)
                break
            
    return adresler

# Loop through all PDF files in the current directory and its subdirectories
for dirpath, dirnames, filenames in os.walk(top_level_folder):
    for filename in filenames:
        if filename.endswith('.pdf') or filename.endswith('.PDF'):
            # Open the PDF file
            file_path = os.path.join(dirpath, filename)

            with fitz.open(file_path) as pdf_file:
                # Check if the PDF file has at least two pages
                if pdf_file.page_count >= 2:
                    # Check if the second page contains the target text
                    for page_number, file in enumerate(pdf_file):
                                                                           
                        if "DEĞERLENDİRME FORMU" in file.get_text():
                            pdf_dirname = os.path.dirname(file_path)
                            file_name = os.path.splitext(os.path.basename(pdf_dirname))[0]
                            site_no = str(re.findall('\d+', file_name)[0])
                            site_name = str(re.findall(r'-(\w+)', file_name)[0] )                          
                            site_id = site_no +'/'+ str(site_name)
                            site_no=int(site_no)
                            pdf_reader = PdfReader(file_path)
                            
                                                   
                            try:
                                if page_number: 
                                    tables = tabula.read_pdf(file_path, pages=str(page_number))
                                    if not tables:
                                        print("No tables found on the page.")
                                    else:
                                        first_table = tables[0]  # Assuming the value you want is in the first table
                                        site_name = first_table.loc[0]
                           
                            except(KeyError):
                                continue
                                                                                                                                                                    
                            site_address="" 
                            latitudes = ''
                            longitudes = ''
       
                            for index, line in enumerate(pdf_file[page_number].get_text().splitlines()):
                                adresler = parse_adres(line)

                                if adresler:
                                    for adres in adresler:
                                        site_address = site_address +' '+adres
                                if '°' in line:
                                    if 'N' in line:
                                        latitudes = line.strip()
                                    elif 'E' in line:
                                        longitudes = line.strip()                                        
                                        
                                                                                                           # except (IndexError, KeyError):
                            #     print("Address extraction failed for file:", file_path)
  
                            # Concatenate a new row to the DataFrame
                            new_row = pd.DataFrame({
                                "File Name": [file_name],
                                "Address": [site_address],
                                "Coordinates": [latitudes + ' ' + longitudes],
                                "Site No": [site_no],
                                "Site ID":[site_id],
                                "Site Name":[site_name]
                            })
                            df = pd.concat([df, new_row], ignore_index=True)

# Write the DataFrame to an Excel file
output_dir = os.path.join(r"C:\Users\T460s\Desktop\testPDf\PDFs", "OUTPUTs")  # Path to output directory
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
output_file = os.path.join(output_dir, "TK_output.xlsx")
df.to_excel(output_file, index=False)
