import os
import glob

import pandas as pd

from pdfminer3.layout import LAParams, LTTextBox
from pdfminer3.pdfpage import PDFPage
from pdfminer3.pdfinterp import PDFResourceManager
from pdfminer3.pdfinterp import PDFPageInterpreter
from pdfminer3.converter import PDFPageAggregator
from pdfminer3.converter import TextConverter
import io

import re


def ex_data(f):
    resource_manager = PDFResourceManager()
    fake_file_handle = io.StringIO()
    converter = TextConverter(resource_manager, fake_file_handle, laparams=LAParams())
    page_interpreter = PDFPageInterpreter(resource_manager, converter)

    # with open('./Call_Info-WO-025046986.pdf', 'rb') as fh:
    with open(f, 'rb') as fh:

        for page in PDFPage.get_pages(fh,
                                      caching=True,
                                      check_extractable=True):
            page_interpreter.process_page(page)

        text = fake_file_handle.getvalue()

# close open handles
    converter.close()
    fake_file_handle.close()

    # print(type(text))

    def split_string_on_newline(input_string):
        lines = input_string.split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        return lines

    input_string = text
    lines = split_string_on_newline(input_string)
# print(lines)

    def extract_fields(data_list, start_string, end_string):
        # Find the index of the start string
        start_index = data_list.index(start_string) + 1
        # Find the index of the end string
        end_index = data_list.index(end_string)
        # Extract the fields between the start and end indices
        fields = data_list[start_index:end_index]
        return fields


    start_string = 'Case Number:'
    end_string = 'Overall service experience rating for this case:'


    fields = extract_fields(lines, start_string, end_string)
# print("Extracted fields:")
# for field in fields:
        # print(field)

    def extract_dates(data_list):
        dates = []
        # Regular expression pattern to match dates in the format mm/dd/yyyy
        date_pattern = r'\b\d{1,2}/\d{1,2}/\d{4}\b'

        # Iterate over the list elements
        for item in data_list:
            # Find all dates in the current item
            matches = re.findall(date_pattern, item)
            # Add all found dates to the dates list
            dates.extend(matches)

        return dates

    required_data = extract_dates(fields)[:3]
    case_id = fields[0].split('/')[0]
    required_data.append(case_id)

    # print(required_data)
    return required_data

# Get the current directory
current_directory = os.getcwd()

# Search for all PDF files in the current directory
pdf_files = glob.glob(os.path.join(current_directory, '*.pdf'))

final_data = []

# Print the names of the PDF files
for i, pdf_file in enumerate(pdf_files):
    # print(i+1)
    final_data.append(ex_data(pdf_file))
    # print(os.path.basename(pdf_file))

# for i, j in enumerate(final_data):
    # print(i+1,j)

df = pd.DataFrame(final_data)

print(df)

excel_file = 'data.xlsx'

df.to_excel(excel_file, index=False)
print("Done !")
