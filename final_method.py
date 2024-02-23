from PyPDF2 import PdfReader
import pandas as pd
import re

def pdf_to_excel_formating(input_path, output_path, extraction_method, format_in, word_divider, number_of_line, words_per_line=5, empty_lines_input=10, each_word_in='y', custom_entry='lorem'):
    def open_pdf(input_path):
        text=''
        reader=PdfReader(input_path)
        for page in reader.pages:
            text+=page.extract_text()
        return text
    
    if extraction_method=='column' or extraction_method=='row' or extraction_method=='space' or extraction_method=='word':    
        text=open_pdf(input_path)    
        words = re.findall(r'\S+', text)
        segments = [' '.join(words[i:i+word_divider]) for i in range(0, len(words), word_divider)]
        segment_count = len(segments) // number_of_line
        if format_in=='column format':
            data = {'Column{}'.format(i+1): segments[i:i+segment_count] for i in range(number_of_line)}
            df = pd.DataFrame(data)
            df.to_excel(output_path, index=False)
            print(df)
        elif format_in=='row format':
            data = {'Column{}'.format(i+1): segments[i:i+segment_count] for i in range(number_of_line)}
            df_row= pd.DataFrame(data)
            row_format=df_row.transpose()
            row_format.to_excel(output_path, index=False)
            print(row_format)
        print("Excel file 'formatted_output.xlsx' has been created successfully.")
    
    elif extraction_method=='line':
        def extract_text_from_pdf(pdf_file_path):
            text = ""
            with open(pdf_file_path, 'rb') as file:
                pdf_reader = PdfReader(file)
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

        extracted_text = extract_text_from_pdf(input_path)

        formatted_text = format_text(extracted_text, words_per_line, empty_lines_input)

        rows = formatted_text.strip().split('\n')

        if each_word_in == 'y':
            df = pd.DataFrame([row.split() for row in rows])
            print(df)
        else:
            df = pd.DataFrame([row for row in rows])
            print(df)

        df.to_excel(output_path, index=False, header=False)

    elif extraction_method=='exact':
        reader = PdfReader(input_path)
        all_text = []
        for page in reader.pages:
            text = page.extract_text()
            all_text.extend(text.split('\n'))

        data = {'Text': all_text}
        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False)
        print(df)

    elif extraction_method == 'custom':
        reader = PdfReader(input_path)
        anything = custom_entry

        output_data = []

        for page in reader.pages:
            text = page.extract_text().lower()
            lines = text.split('\n') 

            for line in lines:
                match_indices = [m.start() for m in re.finditer(re.escape(anything), line)]
                if match_indices:
                    parts = []
                    last_index = 0
                    for index in match_indices:
                        parts.append(line[last_index:index])
                        parts.append(anything)
                        last_index = index + len(anything)
                    parts.append(line[last_index:])
                    output_data.extend(parts)
                else:
                    output_data.append(line)
        if format_in=='row format':
            df = pd.DataFrame(output_data, columns=['Text'])
            df.to_excel(output_path, index=False)
            print(df)
        elif format_in == 'column format':
            df = pd.DataFrame(output_data, columns=['Text'])
            df_transpose=df.transpose()
            df_transpose.to_excel(output_path, index=False)
            print(df_transpose)

    print(f"Successfully converted the file in excel by the name of {output_path} \n method of extraction is {extraction_method} \n the excel file is formated in {format_in} \n the words are divided by {word_divider} \n number of { 'row' if format_in == 'row format' else 'column' } are {number_of_line}")

pdf_to_excel_formating(input_path='sample-pdf-file.pdf',output_path='output_sample.xlsx', extraction_method='custom', format_in='column format', word_divider=4,number_of_line=5, words_per_line=16, empty_lines_input=1, each_word_in='n', custom_entry='i')
