from PyPDF2 import PdfReader
import re
import pandas as pd

pdf_file = input("Enter PDF file name: ")
columns = int(input("Enter the number of columns: "))
reader = PdfReader(pdf_file)
text = ""
for page in range(len(reader.pages)):
    text_extracted = reader.pages[page].extract_text().replace('\n', ' ')
    text = text + ' ' + text_extracted

words = re.findall(r'\S+', text)
num_words = len(words)
num_rows = (num_words + columns - 1) // columns

words.extend([''] * (num_rows * columns - num_words))
words_2d = [words[i*columns:(i+1)*columns] for i in range(num_rows)]
data = {f"Column_{i+1}": column for i, column in enumerate(zip(*words_2d))}
df = pd.DataFrame(data)

df.to_excel("output.xlsx", index=False)
print(df)
