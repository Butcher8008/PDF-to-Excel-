import PyPDF2
import pandas as pd
file_name = input("Enter PDF file name: ")
columns = int(input("Enter Number of Columns: "))

reader = PyPDF2.PdfReader(file_name)
combined_text = ''
for page_num in range(len(reader.pages)):
    page_text = reader.pages[page_num].extract_text().replace('\n', '')
    combined_text += page_text

segment_length = len(combined_text) // columns

segments = [combined_text[i:i+segment_length] for i in range(0, len(combined_text), segment_length)]
dfs = [] 
for segment in segments:
    data = {f"Column_{i+1}": segment for i in range(columns)}
    df = pd.DataFrame(data, index=[segment])  # Create a DataFrame with a single row
    dfs.append(df)  # Append each DataFrame to the list
    print(segment)

result_df = pd.concat(dfs, ignore_index=True)  # Concatenate all DataFrames
result_df.to_excel("output.xlsx", index=False)