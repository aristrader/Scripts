import pandas as pd
import xlwt
from xlwt import Workbook

# Load the existing Excel file
input_file = '/Users/swapnilagarwal/Visual_Studio_Projects/Excel/SUNNY.MSC2SEM.CHEMISTRY.2024.xls'  # Replace with your actual file path
output_file = '/Users/swapnilagarwal/Visual_Studio_Projects/Excel/resultstemp/CHEMISTRY_RESULTS.xls'  # Replace with your desired output file path

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
    'Practical_1_Name': 9,
    'Practical_2_Name': 10,
    'Practical_3_Name': 11,
    'Subject_1_Main_Marks': 12,
    'Subject_1_Main_Marks_Suffix': 13,
    'Subject_1_CCE': 14,
    'Subject_1_CCE_Suffix': 15,
    'Subject_1_Total': 16,
    'Subject_1_Total_Suffix': 17,
    'Subject_2_Main_Marks': 18,
    'Subject_2_Main_Marks_Suffix': 19,
    'Subject_2_CCE': 20,
    'Subject_2_CCE_Suffix': 21,
    'Subject_2_Total': 22,
    'Subject_2_Total_Suffix': 23,
    'Subject_3_Main_Marks': 24,
    'Subject_3_Main_Marks_Suffix': 25,
    'Subject_3_CCE': 26,
    'Subject_3_CCE_Suffix': 27,
    'Subject_3_Total': 28,
    'Subject_3_Total_Suffix': 29,
    'Subject_4_Main_Marks': 30,
    'Subject_4_Main_Marks_Suffix': 31,
    'Subject_4_CCE': 32,
    'Subject_4_CCE_Suffix': 33,
    'Subject_4_Total': 34,
    'Subject_4_Total_Suffix': 35,
    'Practical_1': 36,
    'Practical_1_Suffix': 37,
    'Practical_2': 38,
    'Practical_2_Suffix': 39,
    'Practical_3': 40,
    'Practical_3_Suffix': 41,
    'Therory_Total': 42,
    'Out_Of_Theory': 43,
    'Practical_Total': 44,
    'Out_Of_Practical': 45,
    'SEM_TOTAL': 46,
    'OUT_OF_SEM': 47,
    'RESULT': 48,
    'CATEGORY': 49
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
def split_subject_1(subject, char_limit=30):
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
    # subject_4_first, subject_4_second, subject_4_third = split_subject_4(row.iloc[columns['Subject_4']])
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
        # # Row 2 (split row for subject if needed)
        # [
        #     '', '', row.iloc[columns['Father_Name']], '    ' + subject_1_second if subject_1_second else '', '', '', '', '', '', '',
        #     '', '', ''
        # ],
        # Row 3
        [
            '', '', 
            row.iloc[columns['Father_Name']], 
            subject_2_first, '85', '29', '15', '05', '100', '34',
            (str(row.iloc[columns['Subject_2_Main_Marks']]) if not pd.isna(row.iloc[columns['Subject_2_Main_Marks']]) else '') + 
            (str(row.iloc[columns['Subject_2_Main_Marks_Suffix']]) if not pd.isna(row.iloc[columns['Subject_2_Main_Marks_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_2_CCE']]) if not pd.isna(row.iloc[columns['Subject_2_CCE']]) else '') + 
            (str(row.iloc[columns['Subject_2_CCE_Suffix']]) if not pd.isna(row.iloc[columns['Subject_2_CCE_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_2_Total']]) if not pd.isna(row.iloc[columns['Subject_2_Total']]) else '') + 
            (str(row.iloc[columns['Subject_2_Total_Suffix']]) if not pd.isna(row.iloc[columns['Subject_2_Total_Suffix']]) else '')
        ],
        # # Row 4 (split row for subject if needed)
        # [
        #     '', row.iloc[columns['Enrollement_NO']], '', '    ' + subject_2_second if subject_2_second else '', '', '', '', '', '', '',
        #     '', '', ''
        # ],
        # Row 5
        [
            '', '', row.iloc[columns['Mother_Name']], 
            subject_3_first, 
            '85', '29', '15', '05', '100', '34',
            (str(row.iloc[columns['Subject_3_Main_Marks']]) if not pd.isna(row.iloc[columns['Subject_3_Main_Marks']]) else '') + 
            (str(row.iloc[columns['Subject_3_Main_Marks_Suffix']]) if not pd.isna(row.iloc[columns['Subject_3_Main_Marks_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_3_CCE']]) if not pd.isna(row.iloc[columns['Subject_3_CCE']]) else '') + 
            (str(row.iloc[columns['Subject_3_CCE_Suffix']]) if not pd.isna(row.iloc[columns['Subject_3_CCE_Suffix']]) else ''),

            (str(row.iloc[columns['Subject_3_Total']]) if not pd.isna(row.iloc[columns['Subject_3_Total']]) else '') + 
            (str(row.iloc[columns['Subject_3_Total_Suffix']]) if not pd.isna(row.iloc[columns['Subject_3_Total_Suffix']]) else '')
        ],
        # # Row 6 (split row for subject if needed)
        # [
        #     '', '', '', '    ' + subject_3_second if subject_3_second else '', '', '', '', '', '', '',
        #     '', '', ''
        # ],
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
        #  # Row 8 (split row for subject if needed)
        # [
        #     '', '', '', '    ' + subject_4_third if subject_4_third else '', '', '', '', '', '', '',
        #     '', '', ''
        # ],
        # Row 10 (split row for subject if needed)
        [
            '', '', '', '', '', '', '', '', '', '',
            '', '', ''
        ],
        # Row 10 (split row for subject if needed)
        [
            '', '', '', '    PRACTICAL:', '', '', '', '', '', '',
            '', '', ''
        ],
        # Row 11
        [
            '', 
            '', 
            '', 
            row.iloc[columns['Practical_1_Name']], 
            '', '', '', '', '66', '26',
            (str(row.iloc[columns['Practical_1']]) if not pd.isna(row.iloc[columns['Practical_1']]) else '') + 
            (str(row.iloc[columns['Practical_1_Suffix']]) if not pd.isna(row.iloc[columns['Practical_1_Suffix']]) else ''),
            '',
            (str(row.iloc[columns['Practical_1']]) if not pd.isna(row.iloc[columns['Practical_1']]) else '') + 
            (str(row.iloc[columns['Practical_1_Suffix']]) if not pd.isna(row.iloc[columns['Practical_1_Suffix']]) else '')
        ],
        # Row 12
        [
            '', 
            '', 
            '', 
            row.iloc[columns['Practical_2_Name']], 
            '', '', '', '', '66', '26',
            (str(row.iloc[columns['Practical_2']]) if not pd.isna(row.iloc[columns['Practical_2']]) else '') + 
            (str(row.iloc[columns['Practical_2_Suffix']]) if not pd.isna(row.iloc[columns['Practical_2_Suffix']]) else ''),
            '',
            (str(row.iloc[columns['Practical_2']]) if not pd.isna(row.iloc[columns['Practical_2']]) else '') + 
            (str(row.iloc[columns['Practical_2_Suffix']]) if not pd.isna(row.iloc[columns['Practical_2_Suffix']]) else '')
        ],
        # Row 12
        [
            '', 
            '', 
            '', 
            row.iloc[columns['Practical_3_Name']], 
            '', '', '', '', '68', '28',
            (str(row.iloc[columns['Practical_3']]) if not pd.isna(row.iloc[columns['Practical_3']]) else '') + 
            (str(row.iloc[columns['Practical_3_Suffix']]) if not pd.isna(row.iloc[columns['Practical_3_Suffix']]) else ''),
            '',
            (str(row.iloc[columns['Practical_3']]) if not pd.isna(row.iloc[columns['Practical_3']]) else '') + 
            (str(row.iloc[columns['Practical_3_Suffix']]) if not pd.isna(row.iloc[columns['Practical_3_Suffix']]) else '')
        ],
        # Row 13
        [
            '     CATEG.:' + (row.iloc[columns['CATEGORY']] if not pd.isna(row.iloc[columns['CATEGORY']]) else ''),
            '', 
            '    THEORY MARKS:' + str(row.iloc[columns['Therory_Total']])+ '/' + str(row.iloc[columns['Out_Of_Theory']]),
            '    PRACTICAL MARKS:' + str(row.iloc[columns['Practical_Total']])+ '/' + str(row.iloc[columns['Out_Of_Practical']]), 
            'TOTAL MARKS:' + str(row.iloc[columns['SEM_TOTAL']])+ '/' + str(row.iloc[columns['OUT_OF_SEM']]), 
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
