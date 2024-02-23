from PyPDF2 import PdfReader
import re
import pandas as pd

pdf_file = input("Enter PDF file name: ")
custom_word = int(input("Enter custom word: "))
columns = int(input("Enter number of columns: "))
line_break = '\n' * columns

reader = PdfReader(pdf_file)
text = ''
word_segment = ''
current_segment = []

for page in range(len(reader.pages)):
    text_extracted = reader.pages[page].extract_text()
    text += text_extracted

words = re.findall(r'\S+', text)
segment_length = len(words) // columns
for word in words:
    word_segment += word + ' '
    if len(word_segment.split()) == custom_word:
        current_segment.append([segment.strip() for segment in word_segment.split(line_break)])
        segments = [current_segment[i:i+segment_length] for i in range(0, len(current_segment), segment_length)]
        word_segment=''
     
dfs = [] 
for segment in current_segment:
    data = {f"Column_{i+1}": segment for i in range(columns)}
    df = pd.DataFrame(data, index=[segment])  # Create a DataFrame with a single row
    dfs.append(df)  # Append each DataFrame to the list
    print(segment)
    
result_df = pd.concat(dfs, ignore_index=True)  # Concatenate all DataFrames
result_df.to_excel("output.xlsx", index=False)
# print(word_segment, current_segment)
