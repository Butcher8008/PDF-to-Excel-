import pandas as pd
from PyPDF2 import PdfReader
import re

def pdf_to_excel_new(input_path, output_path, extraction_method, custom_break, column_input, save_method, sheet_name='Sheet1'):
    def extract_text(pdf_path):
        reader = PdfReader(pdf_path)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
        return text

    if extraction_method == 'line':
        text = extract_text(input_path)
        lines = text.split('\n')
        new_lines = []
        count = 0
        current_line = ''

        for line in lines:
            if line.strip():
                current_line += line.strip() + ' '
                count += 1
            if count % custom_break == 0:
                new_lines.append(current_line.strip())
                current_line = ''

        if current_line:
            new_lines.append(current_line.strip())

        df = pd.DataFrame(new_lines, columns=['Text'])

    # elif extraction_method == 'word':
    #     text = extract_text(input_path)
    #     words = re.findall(r'\S+', text)
    #     new_segments = []
    #     current_segment = ''

    #     for word in words:
    #         current_segment += word + ' '
    #         if len(current_segment.split()) == custom_break:
    #             new_segments.append(current_segment.strip())
    #             current_segment = ''

    #     if current_segment:
    #         new_segments.append(current_segment.strip())

    #     df = pd.DataFrame(new_segments, columns=['Word'])

    elif extraction_method == 'space':
        line_break = '\n' * column_input
        reader = PdfReader(input_path)
        text = ''
        word_segment = ''
        current_segment = []

        for page in range(len(reader.pages)):
            text_extracted = reader.pages[page].extract_text()
            text += text_extracted

        words = re.findall(r'\S+', text)
        segment_length = len(words) // column_input
        for word in words:
            word_segment += word + ' '
            if len(word_segment.split()) == custom_break:
                current_segment.append([segment.strip() for segment in word_segment.split(line_break)])
                segments = [current_segment[i:i+segment_length] for i in range(0, len(current_segment), segment_length)]
                word_segment=''

        dfs = [] 
        for segment in current_segment:
            data = {f"Column_{i+1}": segment for i in range(column_input)}
            df = pd.DataFrame(data, index=[segment])  # Create a DataFrame with a single row
            dfs.append(df)  # Append each DataFrame to the list
            print(segment)
            
        df = pd.concat(dfs, ignore_index=True)  # Concatenate all DataFrames


    elif extraction_method == 'column':
        columns = int(input("Enter Number of Columns: "))

        reader = PdfReader(input_path)
        combined_text = ''
        for page_num in range(len(reader.pages)):
            page_text = reader.pages[page_num].extract_text().replace('\n', '')
            combined_text += page_text

        segment_length = len(combined_text) // columns

        segments = [combined_text[i:i+segment_length] for i in range(0, len(combined_text), segment_length)]
        dfs = [] 
        for segment in segments:
            data = {f"Column": segment}
            df = pd.DataFrame(data, index=[segment])  # Create a DataFrame with a single row
            dfs.append(df)  # Append each DataFrame to the list
            print(segment)

        df = pd.concat(dfs, ignore_index=True)  # Concatenate all DataFrames

    elif extraction_method=='word':
            line_break = '\n' * column_input
            reader = PdfReader(input_path)
            text = ''
            word_segment = ''
            current_segment = []

            for page in range(len(reader.pages)):
                text_extracted = reader.pages[page].extract_text()
                text += text_extracted

            words = re.findall(r'\S+', text)
            # segment_length = len(words) // column_input
            for word in words:
                word_segment += word + ' '
                if len(word_segment.split()) == custom_break:
                    current_segment.append([segment.strip() for segment in word_segment.split(line_break)])
                    # segments = [current_segment[i:i+segment_length] for i in range(0, len(current_segment), segment_length)]
                    word_segment=''
            print(current_segment)
            dfs = [] 
            count=0
            for segment in current_segment:
                if save_method=='column':
                    data = {f"Column{i+1} ": segment for i in range(column_input) }
                    df = pd.DataFrame(data, index=[segment])  # Create a DataFrame with a single row
                    dfs.append(df)  # Append each DataFrame to the list
                    print(segment)
                else:
                    data = {"Column ": segment  }
                    df = pd.DataFrame(data, index=[segment])  # Create a DataFrame with a single row
                    dfs.append(df)  # Append each DataFrame to the list
                    print(segment)

            df = pd.concat(dfs, ignore_index=True)  # Concatenate all DataFrames


    else:
        pdfReader = PdfReader(input_path)
        pages = len(pdfReader.pages)
        character = extraction_method
        result_list = []

        for i in range(pages):
            read = pdfReader.pages[i]
            text = read.extract_text().split(" ")

            if character in text:
                new_read = read.extract_text().split(character)
                result_list.extend(new_read)
            else:
                for word in text:
                    result_list.extend(list(word))

        df = pd.DataFrame(result_list, columns=[f"Page_{i + 1}" for i in range(pages)])

    if save_method.lower() == 'row':
        df = df.T  # Transpose the DataFrame to switch rows and columns

    df.to_excel(output_path, index=False, sheet_name=sheet_name)
    print(f"PDF converted to Excel successfully. Save method: {extraction_method}{custom_break}{save_method}")

# Example usage:
pdf_to_excel_new("sample-pdf-file.pdf", "output.xlsx", extraction_method='line', custom_break=10, column_input=3, save_method='column', sheet_name='Sheet1')

