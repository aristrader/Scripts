import pandas as pd
import xlwt
from xlwt import Workbook

# Load the existing Excel file
input_file = '/Users/swapnilagarwal/Visual_Studio_Projects/Excel/SUNNY.MSC2SEM.MATHS.2024.xls'  # Replace with your actual file path
output_file = '/Users/swapnilagarwal/Visual_Studio_Projects/Excel/resultstemp/MATHS_RESULTS.xls'  # Replace with your desired output file path

# Read the data from the first sheet of the Excel file
df = pd.read_excel(input_file, engine='xlrd')

# Mapping of column indices
columns = {
    'Roll_NO': 0,
    'Enrollement_NO': 1,
    'Student_Name': 2,
    'Father_Name': 3,
    'Mother_Name': 4,
    'Subject_1': 5,
    'Subject_2': 6,
    'Subject_3': 7,
    'Subject_4': 8,
    'Subject_5': 9,
    'Subject_1_Main_Marks': 10,
    'Subject_1_Main_Marks_Suffix': 11,
    'Subject_1_CCE': 12,
    'Subject_1_CCE_Suffix': 13,
    'Subject_1_Total': 14,
    'Subject_1_Total_Suffix': 15,
    'Subject_2_Main_Marks': 16,
    'Subject_2_Main_Marks_Suffix': 17,
    'Subject_2_CCE': 18,
    'Subject_2_CCE_Suffix': 19,
    'Subject_2_Total': 20,
    'Subject_2_Total_Suffix': 21,
    'Subject_3_Main_Marks': 22,
    'Subject_3_Main_Marks_Suffix': 23,
    'Subject_3_CCE': 24,
    'Subject_3_CCE_Suffix': 25,
    'Subject_3_Total': 26,
    'Subject_3_Total_Suffix': 27,
    'Subject_4_Main_Marks': 28,
    'Subject_4_Main_Marks_Suffix': 29,
    'Subject_4_CCE': 30,
    'Subject_4_CCE_Suffix': 31,
    'Subject_4_Total': 32,
    'Subject_4_Total_Suffix': 33,
    'Subject_5_Total': 34,
    'Subject_5_Total_Suffix': 35,
    'SEM_TOTAL': 36,
    'OUT_OF_SEM': 37,
    'RESULT': 38,
    'CATEGORY': 39
}

transformed_data = []

# Create a new workbook and add a sheet
wb = Workbook()
ws = wb.add_sheet('Sheet1')

#  THIS HERE WE MIGHT NEED TO CENTER ALLIGN ALL COLUMNS INSTEAD OF USING LINE BREAK ALONT WITH THE TEXT WRAPPING ENABLED

# Set column headers
headers = ['S.NO', 'ROLL NO.', "STUDENT's NAME", 'SUBJECTS OFFERED', '    PAPER', '', '    C.C.E', '', 'MAX', 'MIN', 'MARKS', 'OBT', 'TOTAL']
headers2 = ['', 'ENRL.NO.', "FATHER & MOTHER NAME", '', 'MAX', 'MIN', 'MAX', 'MIN', 'MKS', 'MKS', 'PAPER', 'CCE', '',]

# Write headers to the first row of the worksheet
for col_idx, header in enumerate(headers):
    ws.write(0, col_idx, header)
    
# Write headers to the first row of the worksheet
for col_idx, header in enumerate(headers2):
    ws.write(1, col_idx, header)
    
# Function to split the subject name
def split_subject_1(subject, char_limit=35):
    if len(subject) <= char_limit:
        return subject, ''
    # Find the last space within the char limit
    split_idx = subject[:char_limit].rfind(' ')
    if split_idx == -1:  # No space found, break at char_limit
        split_idx = char_limit
    return subject[:split_idx], subject[split_idx+1:]

def split_subject_2(subject, char_limit=30):
    if len(subject) <= char_limit:
        return subject, ''
    # Find the last space within the char limit
    split_idx = subject[:char_limit].rfind(' ')
    if split_idx == -1:  # No space found, break at char_limit
        split_idx = char_limit
    return subject[:split_idx], subject[split_idx+1:]

def split_subject_3(subject, char_limit=30):
    if len(subject) <= char_limit:
        return subject, ''
    # Find the last space within the char limit
    split_idx = subject[:char_limit].rfind(' ')
    if split_idx == -1:  # No space found, break at char_limit
        split_idx = char_limit
    return subject[:split_idx], subject[split_idx+1:]

def split_subject_4(subject, char_limit=30):
    if len(subject) <= char_limit:
        return subject, ''
    # Find the last space within the char limit
    split_idx = subject[:char_limit].rfind(' ')
    if split_idx == -1:  # No space found, break at char_limit
        split_idx = char_limit
    return subject[:split_idx], subject[split_idx+1:]

def split_subject_5(subject, char_limit=30):
    if len(subject) <= char_limit:
        return subject, ''
    # Find the last space within the char limit
    split_idx = subject[:char_limit].rfind(' ')
    if split_idx == -1:  # No space found, break at char_limit
        split_idx = char_limit
    return subject[:split_idx], subject[split_idx+1:]

# def split_subject_4(subject, char_limit=25):
#     # If the subject fits within the first line
#     if len(subject) <= char_limit:
#         return subject, '', ''
    
#     # Find the split index for the first line
#     split_idx_1 = subject[:char_limit].rfind(' ')
#     if split_idx_1 == -1:  # No space found, break at char_limit
#         split_idx_1 = char_limit
    
#     # For the remaining string, find the split for the second line
#     remaining_subject = subject[split_idx_1+1:]
#     if len(remaining_subject) <= char_limit:
#         return subject[:split_idx_1], remaining_subject, ''
    
#     split_idx_2 = remaining_subject[:char_limit].rfind(' ')
#     if split_idx_2 == -1:  # No space found, break at char_limit
#         split_idx_2 = char_limit
    
#     return (
#         subject[:split_idx_1], 
#         remaining_subject[:split_idx_2], 
#         remaining_subject[split_idx_2+1:]
#     )

# Process each row in the DataFrame
for index, row in df.iterrows():
    subject_1_first, subject_1_second = split_subject_1(row.iloc[columns['Subject_1']])
    subject_2_first, subject_2_second = split_subject_2(row.iloc[columns['Subject_2']])
    subject_3_first, subject_3_second = split_subject_3(row.iloc[columns['Subject_3']])
    subject_4_first, subject_4_second = split_subject_4(row.iloc[columns['Subject_4']])
    subject_5_first, subject_5_second = split_subject_5(row.iloc[columns['Subject_5']])
    transformed_row = [
        # Row 1
        [
            index + 1, 
            row.iloc[columns['Roll_NO']], 
            row.iloc[columns['Student_Name']], 
            subject_1_first, 
            '85', '29', '15', '05', '100', '34',
            (str(row.iloc[columns['Subject_1_Main_Marks']]) if not pd.isna(row.iloc[columns['Subject_1_Main_Marks']]) else '') + 
            (str(row.iloc[columns['Subject_1_Main_Marks_Suffix']]) if not pd.isna(row.iloc[columns['Subject_1_Main_Marks_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_1_CCE']]) if not pd.isna(row.iloc[columns['Subject_1_CCE']]) else '') + 
            (str(row.iloc[columns['Subject_1_CCE_Suffix']]) if not pd.isna(row.iloc[columns['Subject_1_CCE_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_1_Total']]) if not pd.isna(row.iloc[columns['Subject_1_Total']]) else '') + 
            (str(row.iloc[columns['Subject_1_Total_Suffix']]) if not pd.isna(row.iloc[columns['Subject_1_Total_Suffix']]) else '')
        ],
        # Row 2 (split row for subject if needed)
        [
            '', '', row.iloc[columns['Father_Name']], '    ' + subject_1_second if subject_1_second else '', '', '', '', '', '', '',
            '', '', ''
        ],
        # Row 3
        [
            '', '', 
            row.iloc[columns['Mother_Name']], 
            subject_2_first, '85', '29', '15', '05', '100', '34',
            (str(row.iloc[columns['Subject_2_Main_Marks']]) if not pd.isna(row.iloc[columns['Subject_2_Main_Marks']]) else '') + 
            (str(row.iloc[columns['Subject_2_Main_Marks_Suffix']]) if not pd.isna(row.iloc[columns['Subject_2_Main_Marks_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_2_CCE']]) if not pd.isna(row.iloc[columns['Subject_2_CCE']]) else '') + 
            (str(row.iloc[columns['Subject_2_CCE_Suffix']]) if not pd.isna(row.iloc[columns['Subject_2_CCE_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_2_Total']]) if not pd.isna(row.iloc[columns['Subject_2_Total']]) else '') + 
            (str(row.iloc[columns['Subject_2_Total_Suffix']]) if not pd.isna(row.iloc[columns['Subject_2_Total_Suffix']]) else '')
        ],
        # Row 4 (split row for subject if needed)
        [
            '', row.iloc[columns['Enrollement_NO']], '', '    ' + subject_2_second if subject_2_second else '', '', '', '', '', '', '',
            '', '', ''
        ],
        # Row 5
        [
            '', '', '', 
            subject_3_first, 
            '85', '29', '15', '05', '100', '34',
            (str(row.iloc[columns['Subject_3_Main_Marks']]) if not pd.isna(row.iloc[columns['Subject_3_Main_Marks']]) else '') + 
            (str(row.iloc[columns['Subject_3_Main_Marks_Suffix']]) if not pd.isna(row.iloc[columns['Subject_3_Main_Marks_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_3_CCE']]) if not pd.isna(row.iloc[columns['Subject_3_CCE']]) else '') + 
            (str(row.iloc[columns['Subject_3_CCE_Suffix']]) if not pd.isna(row.iloc[columns['Subject_3_CCE_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_3_Total']]) if not pd.isna(row.iloc[columns['Subject_3_Total']]) else '') + 
            (str(row.iloc[columns['Subject_3_Total_Suffix']]) if not pd.isna(row.iloc[columns['Subject_3_Total_Suffix']]) else '')
        ],
        # Row 6 (split row for subject if needed)
        [
            '', '', '', '    ' + subject_3_second if subject_3_second else '', '', '', '', '', '', '',
            '', '', ''
        ],
        # Row 7
        [
            '', '', '', 
            subject_4_first, 
            '85', '29', '15', '05', '100', '34',
            (str(row.iloc[columns['Subject_4_Main_Marks']]) if not pd.isna(row.iloc[columns['Subject_4_Main_Marks']]) else '') + 
            (str(row.iloc[columns['Subject_4_Main_Marks_Suffix']]) if not pd.isna(row.iloc[columns['Subject_4_Main_Marks_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_4_CCE']]) if not pd.isna(row.iloc[columns['Subject_4_CCE']]) else '') + 
            (str(row.iloc[columns['Subject_4_CCE_Suffix']]) if not pd.isna(row.iloc[columns['Subject_4_CCE_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_4_Total']]) if not pd.isna(row.iloc[columns['Subject_4_Total']]) else '') + 
            (str(row.iloc[columns['Subject_4_Total_Suffix']]) if not pd.isna(row.iloc[columns['Subject_4_Total_Suffix']]) else '')
        ],
        # Row 8 (split row for subject if needed)
        [
            '', '', '', '    ' + subject_4_second if subject_4_second else '', '', '', '', '', '', '',
            '', '', ''
        ],
        # Row 7
        [
            '', '', '', 
            subject_5_first, 
            '', '', '', '', '100', '40',
            (str(row.iloc[columns['Subject_5_Total']]) if not pd.isna(row.iloc[columns['Subject_5_Total']]) else '') + 
            (str(row.iloc[columns['Subject_5_Total_Suffix']]) if not pd.isna(row.iloc[columns['Subject_5_Total_Suffix']]) else ''),
            '',
            (str(row.iloc[columns['Subject_5_Total']]) if not pd.isna(row.iloc[columns['Subject_5_Total']]) else '') + 
            (str(row.iloc[columns['Subject_5_Total_Suffix']]) if not pd.isna(row.iloc[columns['Subject_5_Total_Suffix']]) else ''),
        ],
        # Row 8 (split row for subject if needed)
        [
            '', '', '', '    ' + subject_5_second if subject_5_second else '', '', '', '', '', '', '',
            '', '', ''
        ],
        # Row 13
        [
            '     CATEG.:' + (row.iloc[columns['CATEGORY']] if not pd.isna(row.iloc[columns['CATEGORY']]) else ''),
            '', 
            '',
            '                                              TOTAL MARKS: ' + str(row.iloc[columns['SEM_TOTAL']])+ '/' + str(row.iloc[columns['OUT_OF_SEM']]), 
            '', 
            '', '', '', 
            'RESULT:' + (row.iloc[columns['RESULT']] if not pd.isna(row.iloc[columns['RESULT']]) else ''), '',  '',
            'R.NO.:' + (str(row.iloc[columns['Roll_NO']]) if not pd.isna(row.iloc[columns['Roll_NO']]) else '')
        ],
        # TERMINATING ROW
        [
            # _
            '-----------------', '------------------------','-----------------------------------------------------',
            '-----------------------------------------------------------------------------------------------------',
            '-------------------', '----------------', '---------------', '---------------', '----------------------', '----------------------', 
            '----------------------', '----------------------', '------------------------'
        ]
    ]
    
    # Replace NaN with empty string in the transformed row
    for sub_row in transformed_row:
        for i in range(len(sub_row)):
            if pd.isna(sub_row[i]):
                sub_row[i] = ''

    # Append the transformed row to the list
    transformed_data.append(transformed_row)

row_idx = 2
for transformed_row in transformed_data:
    for sub_row in transformed_row:
        for col_idx, value in enumerate(sub_row):
            ws.write(row_idx, col_idx, value)
        row_idx += 1

# Adjust column widths
column_widths = {
    0: 256 * 7, 
    1: 256 * 12, 
    2: 256 * 24, 
    3: 256 * 36,
    4: 256 * 6,
    5: 256 * 5,
    6: 256 * 5, 
    7: 256 * 5, 
    8: 256 * 8, 
    9: 256 * 8, 
    10: 256 * 8,
    11: 256 * 8, 
    12: 256 * 10
}

for col, width in column_widths.items():
    ws.col(col).width = width

# Save the workbook to the desired output file path
wb.save(output_file)
