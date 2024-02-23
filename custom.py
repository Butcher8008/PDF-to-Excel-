from PyPDF2 import PdfReader
import re 
import pandas as pd

reader = PdfReader('sample-pdf-file.pdf')
anything = input('Enter any thing: ')

output_data = []

for page in reader.pages:
    text = page.extract_text().lower()
    lines = text.split('\n') 

    for line in lines:
        if re.search(re.escape(anything), line):  
            parts = re.split(re.escape(anything), line) 
            output_data.extend(parts)  
        else:
            output_data.append(line)

df = pd.DataFrame(output_data, columns=['Text'])

df.to_excel('output.xlsx', index=False)
