from PyPDF2 import PdfReader
import pandas as pd

reader = PdfReader('sample-pdf-file.pdf')
all_text = []
for page in reader.pages:
    text = page.extract_text()
    all_text.extend(text.split('\n'))

data = {'Text': all_text}
df = pd.DataFrame(data)
df.to_excel("exact_output.xlsx", index=False)
print(df)
