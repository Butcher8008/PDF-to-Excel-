import PyPDF2
import pandas as pd

def extract_text_from_pdf(pdf_file_path):
    text = ""
    with open(pdf_file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text().replace('\n', ' ')
    return text

def format_text(text, words_per_line, empty_lines_after):
    lines = text.split('\n')
    formatted_text = ''
    word_count = 0
    for line in lines:
        words = line.strip().split()
        for word in words:
            formatted_text += word + ' '
            word_count += 1
            if word_count == words_per_line:
                formatted_text += '\n'
                word_count = 0
                empty_lines_after -= 1
                if empty_lines_after == 0:
                    formatted_text += '\n'
                    empty_lines_after = empty_lines_input
        if word_count == 0:
            formatted_text += '\n'
            empty_lines_after -= 1
            if empty_lines_after == 0:
                formatted_text += '\n'
                empty_lines_after = empty_lines_input
    return formatted_text

# Input parameters
pdf_file_path = 'sample-pdf-file.pdf'
words_per_line = int(input("Enter words per line: "))
empty_lines_input = int(input("Enter number of lines after which you want to have line break: "))

# Extract text from PDF
extracted_text = extract_text_from_pdf(pdf_file_path)

# Format extracted text
formatted_text = format_text(extracted_text, words_per_line, empty_lines_input)

# Split formatted text into rows
rows = formatted_text.strip().split('\n')
each_word_in=input("Enter wether you want each word in seperate row or not : type y for yes and n for no: ")

if each_word_in == 'y':
    df = pd.DataFrame([row.split() for row in rows])
    print(df)
else:
    df = pd.DataFrame([row for row in rows])
    print(df)
excel_file_path = 'formatted_data.xlsx'
df.to_excel(excel_file_path, index=False, header=False)
print(f"Formatted data has been saved to '{excel_file_path}'")
