import pandas as pd
import xlwt
from xlwt import Workbook

# Load the existing Excel file
input_file = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/FINAL_RESULTS.xlsx'  # Replace with your actual file path
output_file = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/TABULATION.xls'  # Replace with your desired output file path

# Read the data from the first sheet of the Excel file
# df = pd.read_excel(input_file, engine='xlrd')
df = pd.read_excel(input_file, engine='openpyxl')

# Mapping of column indices
columns = {
    'Roll_NO': 0,
    'Enrollement_NO': 1,
    'Student_Name': 2,
    'Father_Name': 3,
    'Mother_Name': 4,
    'Foundation_Course': 5,
    'Foundation_Course_Main_Marks': 6,
    'Foundation_Course_CCE_Marks': 7,
    'Foundation_Course_Total_Marks': 8,
    'Foundation_Course_Max_Credits': 9,
    'Foundation_Course_Grade_Letter': 10,
    'Foundation_Course_Grade_Point': 11,
    'Foundation_Course_Earned_Credit': 12,
    'Foundation_Course_Credit_Points': 13,
    'Foundation_Course_Practical': 14,
    'Foundation_Course_Practical_Main_Marks': 15,
    'Foundation_Course_Practical_CCE_Marks': 16,
    'Foundation_Course_Practical_Total_Marks': 17,
    'Foundation_Course_Practical_Max_Credits': 18,
    'Foundation_Course_Practical_Grade_Letter': 19,
    'Foundation_Course_Practical_Grade_Point': 20,
    'Foundation_Course_Practical_Earned_Credit': 21,
    'Foundation_Course_Practical_Credit_Points': 22,
    'Major_1_Subject': 23,
    'Major_1_Main_Marks': 24,
    'Major_1_CCE_Marks': 25,
    'Major_1_Total_Marks': 26,
    'Major_1_Max_Credits': 27,
    'Major_1_Grade_Letter': 28,
    'Major_1_Grade_Point': 29,
    'Major_1_Earned_Credit': 30,
    'Major_1_Credit_Points': 31,
    'Major_1_Practical_Subject': 32,
    'Major_1_Practical_Main_Marks': 33,
    'Major_1_Practical_CCE_Marks': 34,
    'Major_1_Practical_Total_Marks': 35,
    'Major_1_Practical_Max_Credits': 36,
    'Major_1_Practical_Grade_Letter': 37,
    'Major_1_Practical_Grade_Point': 38,
    'Major_1_Practical_Earned_Credit': 39,
    'Major_1_Practical_Credit_Points': 40,
    'Major_2_Subject': 41,
    'Major_2_Main_Marks': 42,
    'Major_2_CCE_Marks': 43,
    'Major_2_Total_Marks': 44,
    'Major_2_Max_Credits': 45,
    'Major_2_Grade_Letter': 46,
    'Major_2_Grade_Point': 47,
    'Major_2_Earned_Credit': 48,
    'Major_2_Credit_Points': 49,
    'Major_2_Practical_Subject': 50,
    'Major_2_Practical_Main_Marks': 51,
    'Major_2_Practical_CCE_Marks': 52,
    'Major_2_Practical_Total_Marks': 53,
    'Major_2_Practical_Max_Credits': 54,
    'Major_2_Practical_Grade_Letter': 55,
    'Major_2_Practical_Grade_Point': 56,
    'Major_2_Practical_Earned_Credit': 57,
    'Major_2_Practical_Credit_Points': 58,
    'Minor_Subject': 59,
    'Minor_Main_Marks': 60,
    'Minor_CCE_Marks': 61,
    'Minor_Total_Marks': 62,
    'Minor_Max_Credits': 63,
    'Minor_Grade_Letter': 64,
    'Minor_Grade_Point': 65,
    'Minor_Earned_Credit': 66,
    'Minor_Credit_Points': 67,
    'Minor_Practical_Subject': 68,
    'Minor_Practical_Main_Marks': 69,
    'Minor_Practical_CCE_Marks': 70,
    'Minor_Practical_Total_Marks': 71,
    'Minor_Practical_Max_Credits': 72,
    'Minor_Practical_Grade_Letter': 73,
    'Minor_Practical_Grade_Point': 74,
    'Minor_Practical_Earned_Credit': 75,
    'Minor_Practical_Credit_Points': 76,
    'Open_Subject': 77,
    'Open_Main_Marks': 78,
    'Open_CCE_Marks': 79,
    'Open_Total_Marks': 80,
    'Open_Max_Credits': 81,
    'Open_Grade_Letter': 82,
    'Open_Grade_Point': 83,
    'Open_Earned_Credit': 84,
    'Open_Credit_Points': 85,
    'Open_Practical_Subject': 86,
    'Open_Practical_Main_Marks': 87,
    'Open_Practical_CCE_Marks': 88,
    'Open_Practical_Total_Marks': 89,
    'Open_Practical_Max_Credits': 90,
    'Open_Practical_Grade_Letter': 91,
    'Open_Practical_Grade_Point': 92,
    'Open_Practical_Earned_Credit': 93,
    'Open_Practical_Credit_Points': 94,
    'VOC_Subject': 95,
    'VOC_Main_Marks': 96,
    'VOC_CCE_Marks': 97,
    'VOC_Total_Marks': 98,
    'VOC_Max_Credits': 99,
    'VOC_Grade_Letter': 100,
    'VOC_Grade_Point': 101,
    'VOC_Earned_Credit': 102,
    'VOC_Credit_Points': 103,
    'VOC_Practical_Subject': 104,
    'VOC_Practical_Main_marks': 105,
    'VOC_Practical_CCE_Marks': 106,
    'VOC_Practical_Total_marks': 107,
    'VOC_Practical_Max_Credits': 108,
    'VOC_Practical_Grade_Letter': 109,
    'VOC_Practical_Grade_Point': 110,
    'VOC_Practical_Earned_Credit': 111,
    'VOC_Practical_Credit_Points': 112,
    'Project/Internship': 113,
    'Project/Internship_Main_marks': 114,
    'Project/Internship_CCE_marks': 115,
    'Project/Internship_Total_marks': 116,
    'Project/Internship_Max_Credits': 117,
    'Project/Internship_Grade_Letter': 118,
    'Project/Internship_Grade_Point': 119,
    'Project/Internship_Earned_Credit': 120,
    'Project/Internship_Credit_Points': 121,
    'Project/Internship_Practical': 122,
    'Project/Internship_Practical_Main_marks': 123,
    'Project/Internship_Practical_CCE_marks': 124,
    'Project/Internship_Practical_Total_marks': 125,
    'Project/Internship_Practical_Max_Credits': 126,
    'Project/Internship_Practical_Grade_Letter': 127,
    'Project/Internship_Practical_Grade_Point': 128,
    'Project/Internship_Practical_Earned_Credit': 129,
    'Project/Internship_Practical_Credit_Points': 130,
    'Total_Max_Credits': 131,
    'Total_Earned_Credits': 132,
    'Total_Credit_Points': 133,
    'Result': 134,
    'Category': 135,
}

transformed_data = []

# Create a new workbook and add a sheet
wb = Workbook()
ws = wb.add_sheet('Sheet1')

header_college = ['', '', '', 'MLB COLLEGE']
for col_idx, header in enumerate(header_college):
    ws.write(0, col_idx, header)
    
header_terminating_row = [
            # _
            '-----------', '------------------','-----------------------------------------',
            '-----------------------------------------------------------------------',
            '------------', '---------', '--------', '--------', '---------------', '---------------', '---------------', '---------------', '-----------------'
        ]
for col_idx, header in enumerate(header_terminating_row):
    ws.write(1, col_idx, header)

# Set column headers
headers = ['S.NO', 'ROLL NO.\nENRL.NO.', "STUDENT's NAME\nFATHER & MOTHER NAME", 'SUBJECT', 'FC', '50\n70', '50\n30', 'TOT\n100', 'MAX\nCREDITS', 'LETTER\nGRADE', 'GRADE\nPOINT', 'EARNED\nCREDIT', 'CR.POINTS\n(C*xGP)',]

# Set the height for the header row (row index 2)
header_row_height = 256 * 3  # Set height in units of 1/256 of a character width
ws.row(2).height_mismatch = True
ws.row(2).height = header_row_height
header_style = xlwt.easyxf('align: wrap on;')

# Write headers to the third row of the worksheet
for col_idx, header in enumerate(headers):
    ws.write(2, col_idx, header, header_style)
    
for col_idx, header in enumerate(header_terminating_row):
    ws.write(3, col_idx, header)

# Process each row in the DataFrame
for index, row in df.iterrows():
    transformed_row = [
        # Row 1
        [
            # S.NO, Roll. No, Student Name , Foundation Subject, All blanks
            index + 1, row.iloc[columns['Roll_NO']], row.iloc[columns['Student_Name']], row.iloc[columns['Foundation_Course']], 
            '', row.iloc[columns['Foundation_Course_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Foundation_Course_CCE_Marks']], row.iloc[columns['Foundation_Course_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Foundation_Course_Max_Credits']], row.iloc[columns['Foundation_Course_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Foundation_Course_Grade_Point']], row.iloc[columns['Foundation_Course_Earned_Credit']],
            # cr_points
            row.iloc[columns['Foundation_Course_Credit_Points']]
        ],
        # Row 2
        [
            # _, _, Father_name, Course1, main,
            '', '',row.iloc[columns['Father_Name']], row.iloc[columns['Foundation_Course_Practical']], '', row.iloc[columns['Foundation_Course_Practical_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Foundation_Course_Practical_CCE_Marks']], row.iloc[columns['Foundation_Course_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Foundation_Course_Practical_Max_Credits']], row.iloc[columns['Foundation_Course_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Foundation_Course_Practical_Grade_Point']], row.iloc[columns['Foundation_Course_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Foundation_Course_Practical_Credit_Points']]
        ],
        # Row 3
        [
            # _, _, Mother_name, Course2,
            '', '',row.iloc[columns['Mother_Name']], row.iloc[columns['Major_1_Subject']], '', row.iloc[columns['Major_1_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Major_1_CCE_Marks']], row.iloc[columns['Major_1_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Major_1_Max_Credits']], row.iloc[columns['Major_1_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Major_1_Grade_Point']], row.iloc[columns['Major_1_Earned_Credit']],
            # cr_points
            row.iloc[columns['Major_1_Credit_Points']]
        ],
        # Row 6
        [
            # _, _, _, Major_1_subject, main,
            '', '', row.iloc[columns['Enrollement_NO']], row.iloc[columns['Major_1_Practical_Subject']], '', row.iloc[columns['Major_1_Practical_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Major_1_Practical_CCE_Marks']], row.iloc[columns['Major_1_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Major_1_Practical_Max_Credits']], row.iloc[columns['Major_1_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Major_1_Practical_Grade_Point']], row.iloc[columns['Major_1_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Major_1_Practical_Credit_Points']]
        ],
        # Row 7
        [
            # _, _, _, Major_2_subject, main,
            '', '', '', row.iloc[columns['Major_2_Subject']], '', row.iloc[columns['Major_2_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Major_2_CCE_Marks']], row.iloc[columns['Major_2_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Major_2_Max_Credits']], row.iloc[columns['Major_2_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Major_2_Grade_Point']], row.iloc[columns['Major_2_Earned_Credit']],
            # cr_points
            row.iloc[columns['Major_2_Credit_Points']]
        ],
        # Row 8
        [
            # _, _, _, Major_2_subject, main,
            '', '', '', row.iloc[columns['Major_2_Practical_Subject']], '', row.iloc[columns['Major_2_Practical_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Major_2_Practical_CCE_Marks']], row.iloc[columns['Major_2_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Major_2_Practical_Max_Credits']], row.iloc[columns['Major_2_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Major_2_Practical_Grade_Point']], row.iloc[columns['Major_2_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Major_2_Practical_Credit_Points']]
        ],
        # Row 9
        [
            # _, _, _, Minor_subject, main,
            '', '', '', row.iloc[columns['Minor_Subject']], '', row.iloc[columns['Minor_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Minor_CCE_Marks']], row.iloc[columns['Minor_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Minor_Max_Credits']], row.iloc[columns['Minor_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Minor_Grade_Point']], row.iloc[columns['Minor_Earned_Credit']],
            # cr_points
            row.iloc[columns['Minor_Credit_Points']]
        ],
        # Row 10
        [
            # _, _, _, Minor_subject, main,
            '', '', '', row.iloc[columns['Minor_Practical_Subject']], '', row.iloc[columns['Minor_Practical_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Minor_Practical_CCE_Marks']], row.iloc[columns['Minor_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Minor_Practical_Max_Credits']], row.iloc[columns['Minor_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Minor_Practical_Grade_Point']], row.iloc[columns['Minor_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Minor_Practical_Credit_Points']]
        ],
        # Row 11
        [
            # _, _, _, Open_subject, main,
            '', '', '', row.iloc[columns['Open_Subject']], '', row.iloc[columns['Open_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Open_CCE_Marks']], row.iloc[columns['Open_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Open_Max_Credits']], row.iloc[columns['Open_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Open_Grade_Point']], row.iloc[columns['Open_Earned_Credit']],
            # cr_points
            row.iloc[columns['Open_Credit_Points']]
        ],
        # Row 12
        [
            # _, _, _, Open_subject, main,
            '', '', '', row.iloc[columns['Open_Practical_Subject']], '', row.iloc[columns['Open_Practical_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Open_Practical_CCE_Marks']], row.iloc[columns['Open_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Open_Practical_Max_Credits']], row.iloc[columns['Open_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Open_Practical_Grade_Point']], row.iloc[columns['Open_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Open_Practical_Credit_Points']]
        ],
        # Row 13
        [
            # _, _, _, VOC_subject, main,
            '', '', '', row.iloc[columns['VOC_Subject']], '', row.iloc[columns['VOC_Main_Marks']],
            #  cce, obt,
            '', row.iloc[columns['VOC_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['VOC_Max_Credits']], row.iloc[columns['VOC_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['VOC_Grade_Point']], row.iloc[columns['VOC_Earned_Credit']],
            # cr_points
            row.iloc[columns['VOC_Credit_Points']]
        ],
        # Row 14
        [
            # _, _, _, VOC_subject, main,
            '', '', '', row.iloc[columns['VOC_Practical_Subject']], '', row.iloc[columns['VOC_Practical_Main_marks']],
            #  cce, obt,
            row.iloc[columns['VOC_Practical_CCE_Marks']], row.iloc[columns['VOC_Practical_Total_marks']],
            # max_credits, letter_grade,
            row.iloc[columns['VOC_Practical_Max_Credits']], row.iloc[columns['VOC_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['VOC_Practical_Grade_Point']], row.iloc[columns['VOC_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['VOC_Practical_Credit_Points']]
        ],
        # Row 16
        [
            # _, _, _, Project/Internship_practical, main,
            '', '', '', row.iloc[columns['Project/Internship']], '', row.iloc[columns['Project/Internship_Practical_Main_marks']],
            #  cce, obt,
            '', row.iloc[columns['Project/Internship_Practical_Total_marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Project/Internship_Practical_Max_Credits']], row.iloc[columns['Project/Internship_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Project/Internship_Practical_Grade_Point']], row.iloc[columns['Project/Internship_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Project/Internship_Practical_Credit_Points']]
        ],
        [
            # _, _, _, Project/Internship_practical, main,
            '', '', '', row.iloc[columns['Project/Internship_Practical']], '', row.iloc[columns['Project/Internship_Practical_Main_marks']],
            #  cce, obt,
            '', row.iloc[columns['Project/Internship_Practical_Total_marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Project/Internship_Practical_Max_Credits']], row.iloc[columns['Project/Internship_Practical_Grade_Letter']],
            # grade_point, earned_credit,
            row.iloc[columns['Project/Internship_Practical_Grade_Point']], row.iloc[columns['Project/Internship_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Project/Internship_Practical_Credit_Points']]
        ],
        # Row 17
        [
            # _, _, _, _, 'TOTAL'
            '', '', '', '                                                          TOTAL:', '',
            # _, _, _, total max credits
            '', '', '', row.iloc[columns['Total_Max_Credits']],
            # _, Total earned creadits, total credit points
            '', '', row.iloc[columns['Total_Earned_Credits']], row.iloc[columns['Total_Credit_Points']]
        ],
        # Row 18
        [
            # _, 'AGPA = CREDIT POINTS/MAX CREDIT', _, _, 'AGPA : agpa var'
            '', '', '', '', 
            '',
            # _, _, _
            '', '', '', '',
            # _, _, _, _
            '', '', ''
        ],
        # Row 19
        [
            # _, 'CATEG.: category var', _, _, 'RESULT: result var'
            '', 'CATEG.:' + row.iloc[columns['Category']], '', '                                            RESULT:' + row.iloc[columns['Result']],
            # _, _, _
            '', '', '', '', '', '',
            # _, 'Rol.No.:Roll number var, _,
            '', 'Rol.No.:' + str(row.iloc[columns['Roll_NO']])
        ],
        # TERMINATING ROW
        [
            # _
            '-----------', '------------------','-----------------------------------------',
            '-----------------------------------------------------------------------',
            '------------', '---------', '--------', '--------', '---------------', '---------------', '---------------', '---------------', '-----------------'
        ]
    ]
    
    # Replace NaN with empty string in the transformed row
    for sub_row in transformed_row:
        for i in range(len(sub_row)):
            if pd.isna(sub_row[i]):
                sub_row[i] = ''

    # Append the transformed row to the list
    transformed_data.append(transformed_row)

row_idx = 4
for transformed_row in transformed_data:
    for sub_row in transformed_row:
        for col_idx, value in enumerate(sub_row):
            ws.write(row_idx, col_idx, value)
        row_idx += 1

# Adjust column widths
column_widths = {
    0: 256 * 6, 
    1: 256 * 10, 
    2: 256 * 22, 
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



# We need allignment , text size, more higher text size for the college name, boarders bold for the headers. line before and after.