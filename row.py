from PyPDF2 import PdfReader
import re
import pandas as pd

reader = PdfReader("sample-pdf-file.pdf")
word_count = int(input("Enter the number of words after you want a line break: "))
row_break = '\n' * word_count
rows_allowed = int(input("Enter the rows or either columns you want to add the text in: "))
extraction_method=input("Enter Extraction method : e.g row\column ")

text = ''
for page in range(len(reader.pages)):
    text += reader.pages[page].extract_text().replace('\n', '')

words = re.findall(r'\S+', text)

segments = [' '.join(words[i:i+word_count]) for i in range(0, len(words), word_count)]
segment_count = len(segments) // rows_allowed
if extraction_method=='column':
    data = {'Column{}'.format(i+1): segments[i:i+segment_count] for i in range(rows_allowed)}
    df = pd.DataFrame(data)
    print(df)
    df.to_excel('formatted_output_column.xlsx', index=False)
else:
    data = {'Column{}'.format(i+1): segments[i:i+segment_count] for i in range(rows_allowed)}
    df_row= pd.DataFrame(data)
    row_format=df_row.transpose()
    row_format.to_excel('formatted_output_row.xlsx', index=False)
    print(row_format)
print("Excel file 'formatted_output.xlsx' has been created successfully.")
