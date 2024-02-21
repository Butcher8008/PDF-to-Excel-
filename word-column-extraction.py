from PyPDF2 import PdfReader
import re
import pandas as pd

pdf_file= input("Enter PDF file name: ")
columns=int(input("Enter the number of columns: "))
line_break= '\n' * columns
reader = PdfReader(pdf_file)
text = ""
current_alfaz=''
for page in range(len(reader.pages)):
    text_extracted= reader.pages[page].extract_text().replace('\n','')
    text=text+text_extracted

alfaz_segment= len(text) // columns

word=re.findall(r'\S+', text)
current_alfaz= ' '.join(f"'{w}'" for w in word)

current_alfaz.split(line_break)
dfs=[]
for alfaz in current_alfaz:
    data = {f"Column_{i+1}": alfaz for i in range(columns)}
    df = pd.DataFrame(data, index=[alfaz])
    dfs.append(df)
    print(alfaz)

result_df = pd.concat(dfs, ignore_index=True)  # Concatenate all DataFrames
result_df.to_excel("output.xlsx", index=False)

