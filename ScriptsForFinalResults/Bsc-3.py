import pandas as pd
import xlwt
from xlwt import Workbook

# Load the existing Excel file
input_file = '/Users/swapnilagarwal/Desktop/SUNNEY.BSC3.UNI.2024.xls'  # Replace with your actual file path
output_file = '/Users/swapnilagarwal/Desktop/bsc3results.xls'  # Replace with your desired output file path

# Read the data from the first sheet of the Excel file
df = pd.read_excel(input_file, engine='xlrd')

# Mapping of column indices
columns = {
    'Roll_NO': 0,
    'Enrollement_NO': 1,
    'Student_Name': 2,
    'Father_Name': 3,
    'Mother_Name': 4,
    # 'Foundation_Course': 5,
    'Foundation_Course_Subject_I_A_Marks': 16,
    'Foundation_Course_Subject_I_B_Marks': 17,
    'Foundation_Course_Subject_I_Total_Marks': 18,
    'Foundation_Course_Subject_I_Max_Credits': 19,
    'Foundation_Course_Subject_I_Grade_Letter': 20,
    'Foundation_Course_Subject_I_Grade_Point': 21,
    'Foundation_Course_Subject_I_Earned_Credit': 22,
    'Foundation_Course_Subject_I_Credit_Points': 23,
    'Foundation_Course_Subject_II_A_Marks': 24,
    'Foundation_Course_Subject_II_B_Marks': 25,
    'Foundation_Course_Subject_II_Total_Marks': 26,
    'Foundation_Course_Subject_II_Max_Credits': 27,
    'Foundation_Course_Subject_II_Grade_Letter': 28,
    'Foundation_Course_Subject_II_Grade_Point': 29,
    'Foundation_Course_Subject_II_Earned_Credit': 30,
    'Foundation_Course_Subject_II_Credit_Points': 31,
    'Major_1_Subject': 5,
    'Major_1_Main_Marks': 32,
    'Major_1_CCE_Marks': 34,
    'Major_1_Total_Marks': 35,
    'Major_1_Max_Credits': 36,
    'Major_1_Grade_Letter': 37,
    'Major_1_Grade_Point': 38,
    'Major_1_Earned_Credit': 39,
    'Major_1_Credit_Points': 40,
    'Major_1_Practical_Subject': 6,
    # 'Major_1_Practical_Main_Marks': 32,
    # 'Major_1_Practical_CCE_Marks': 34,
    'Major_1_Practical_Total_Marks': 48,
    'Major_1_Practical_Max_Credits': 49,
    'Major_1_Practical_Grade_Letter': 50,
    'Major_1_Practical_Grade_Point': 51,
    'Major_1_Practical_Earned_Credit': 52,
    'Major_1_Practical_Credit_Points': 53,
    'Major_2_Subject': 7,
    'Major_2_Main_Marks': 33,
    'Major_2_CCE_Marks': 41,
    'Major_2_Total_Marks': 42,
    'Major_2_Max_Credits': 43,
    'Major_2_Grade_Letter': 44,
    'Major_2_Grade_Point': 45,
    'Major_2_Earned_Credit': 46,
    'Major_2_Credit_Points': 47,
    'Major_2_Practical_Subject': 8,
    # 'Major_2_Practical_Main_Marks': 33,
    # 'Major_2_Practical_CCE_Marks': 41,
    'Major_2_Practical_Total_Marks': 54,
    'Major_2_Practical_Max_Credits': 55,
    'Major_2_Practical_Grade_Letter': 56,
    'Major_2_Practical_Grade_Point': 57,
    'Major_2_Practical_Earned_Credit': 58,
    'Major_2_Practical_Credit_Points': 59,
    'Minor_Subject': 9,
    'Minor_Main_Marks': 60,
    'Minor_CCE_Marks': 61,
    'Minor_Total_Marks': 62,
    'Minor_Max_Credits': 63,
    'Minor_Grade_Letter': 64,
    'Minor_Grade_Point': 65,
    'Minor_Earned_Credit': 66,
    'Minor_Credit_Points': 67,
    'Minor_Practical_Subject': 10,
    # 'Minor_Practical_Main_Marks': 60,
    # 'Minor_Practical_CCE_Marks': 61,
    'Minor_Practical_Total_Marks': 68,
    'Minor_Practical_Max_Credits': 69,
    'Minor_Practical_Grade_Letter': 70,
    'Minor_Practical_Grade_Point': 71,
    'Minor_Practical_Earned_Credit': 72,
    'Minor_Practical_Credit_Points': 73,
    'Open_Subject': 11,
    'Open_Main_Marks': 74,
    'Open_CCE_Marks': 75,
    'Open_Total_Marks': 76,
    'Open_Max_Credits': 77,
    'Open_Grade_Letter': 79,
    'Open_Grade_Point': 80,
    'Open_Earned_Credit': 78,
    'Open_Credit_Points': 81,
    'Open_Practical_Subject': 12,
    # 'Open_Practical_Main_Marks': 56,
    # 'Open_Practical_CCE_Marks': 57,
    'Open_Practical_Total_Marks': 82,
    'Open_Practical_Max_Credits': 83,
    'Open_Practical_Grade_Letter': 84,
    'Open_Practical_Grade_Point': 85,
    'Open_Practical_Earned_Credit': 86,
    'Open_Practical_Credit_Points': 87,
    'VOC_Subject': 13,
    'VOC_Main_Marks': 88,
    # 'VOC_CCE_Marks': 89, -> hard code to 0 or null whatever suits score
    'VOC_Total_Marks': 89,
    'VOC_Max_Credits': 90,
    'VOC_Grade_Letter': 91,
    'VOC_Grade_Point': 92,
    'VOC_Earned_Credit': 93,
    'VOC_Credit_Points': 94,
    'VOC_Practical': 14,
    'VOC_Practical_Main_marks': 102,
    'VOC_Practical_Total_marks': 102,
    'VOC_Practical_Max_Credits': 103,
    'VOC_Practical_Grade_Letter': 104,
    'VOC_Practical_Grade_Point': 105,
    'VOC_Practical_Earned_Credit': 106,
    'VOC_Practical_Credit_Points': 107,
    'Project/Internship': 15,
    'Project/Internship_Practical_Main_marks': 95,
    'Project/Internship_Practical_Total_marks': 96,
    'Project/Internship_Practical_Max_Credits': 97,
    'Project/Internship_Practical_Grade_Letter': 98,
    'Project/Internship_Practical_Grade_Point': 99,
    'Project/Internship_Practical_Earned_Credit': 100,
    'Project/Internship_Practical_Credit_Points': 101,
    'Total_Max_Credits': 108,
    'Total_Earned_Credits': 109,
    'Total_Credit_Points': 110,
    'AGPA_1': 111,
    'AGPA_2': 112,
    'AGPA_3': 113,
    # 'AGPA': 91,
    'CGPA': 114,
    'Percentage': 115,
    'Result': 116,
    'Category': 117
}

transformed_data = []

# Create a new workbook and add a sheet
wb = Workbook()
ws = wb.add_sheet('Sheet1')

#  THIS HERE WE MIGHT NEED TO CENTER ALLIGN ALL COLUMNS INSTEAD OF USING LINE BREAK ALONT WITH THE TEXT WRAPPING ENABLED

# Set column headers
headers = ['S.NO', 'ROLL NO.\nENRL.NO.', "STUDENT's NAME\nFATHER & MOTHER NAME", 'SUBJECT', 'FC', '50\n70', '50\n30', 'TOT\n100', 'MAX\nCREDITS', 'LETTER\nGRADE', 'GRADE\nPOINT', 'EARNED\nCREDIT', 'CR.POINTS\n(C*xGP)',]

# Write headers to the first row of the worksheet
for col_idx, header in enumerate(headers):
    ws.write(0, col_idx, header)

# Process each row in the DataFrame
for index, row in df.iterrows():
    transformed_row = [
        # Row 1
        [
            # S.NO, Roll. No, Student Name , Foundation Subject, All blanks
            index + 1, row.iloc[columns['Roll_NO']], row.iloc[columns['Student_Name']], '1. FOUNDATION COURSE', '', '', '', '', '', '', '', ''
        ],
        # Row 2
        [
            # _, _, Father_name, Course1, main,
            '', '',row.iloc[columns['Father_Name']], 'I (A) HINDI LANGUAGE', '', row.iloc[columns['Foundation_Course_Subject_I_A_Marks']],
            #  cce, obt,
            row.iloc[columns['Foundation_Course_Subject_I_B_Marks']], row.iloc[columns['Foundation_Course_Subject_I_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Foundation_Course_Subject_I_Max_Credits']], row.iloc[columns['Foundation_Course_Subject_I_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Foundation_Course_Subject_I_Grade_Point']], row.iloc[columns['Foundation_Course_Subject_I_Earned_Credit']],
            # cr_points
            row.iloc[columns['Foundation_Course_Subject_I_Credit_Points']]
        ],
        # Row 3
        [
            # _, _, Mother_name, Course2,
            '', '',row.iloc[columns['Mother_Name']], 'I (B) ENGLISH LANGUAGE', '', '', '', '', '', '', '', ''
        ],
        # Row 4
        [
            # _, _, _, Course3, main,
            '', '', '', 'II (A) PER.DEVE. & CHAR.BULD.', '', row.iloc[columns['Foundation_Course_Subject_II_A_Marks']],
            #  cce, obt,
            row.iloc[columns['Foundation_Course_Subject_II_B_Marks']], row.iloc[columns['Foundation_Course_Subject_II_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Foundation_Course_Subject_II_Max_Credits']], row.iloc[columns['Foundation_Course_Subject_II_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Foundation_Course_Subject_II_Grade_Point']], row.iloc[columns['Foundation_Course_Subject_II_Earned_Credit']],
            # cr_points
            row.iloc[columns['Foundation_Course_Subject_II_Credit_Points']]
        ],
        # Row 5
        [
            # _, _, _, Course4,
            '', '', '', 'II (B) DIGITAL AWAR.-CYBER SECURITY', '', '', '', '', '', '', '', ''
        ],
        # Row 6
        [
            # _, _, _, Major_1_subject, main,
            '', '', row.iloc[columns['Enrollement_NO']], row.iloc[columns['Major_1_Subject']], '', row.iloc[columns['Major_1_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Major_1_CCE_Marks']], row.iloc[columns['Major_1_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Major_1_Max_Credits']], row.iloc[columns['Major_1_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Major_1_Grade_Point']], row.iloc[columns['Major_1_Earned_Credit']],
            # cr_points
            row.iloc[columns['Major_1_Credit_Points']]
        ],
        # Row 7
        [
            # _, _, _, Major_1_Practical_subject, main, cce
            '', '', '', ('     ' + str(row.iloc[columns['Major_1_Practical_Subject']]) if pd.notna(row.iloc[columns['Major_1_Practical_Subject']]) else ''), '', '', '', #row.iloc[columns['Major_1_Practical_Main_Marks']], row.iloc[columns['Major_1_Practical_CCE_Marks']],
            # obt,
            row.iloc[columns['Major_1_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Major_1_Practical_Max_Credits']], row.iloc[columns['Major_1_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Major_1_Practical_Grade_Point']], row.iloc[columns['Major_1_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Major_1_Practical_Credit_Points']]
        ],
        # Row 8
        [
            # _, _, _, Major_2_subject, main,
            '', '', '', '   ' + row.iloc[columns['Major_2_Subject']], '', row.iloc[columns['Major_2_Main_Marks']],
            #  cce, obt,
            row.iloc[columns['Major_2_CCE_Marks']], row.iloc[columns['Major_2_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Major_2_Max_Credits']], row.iloc[columns['Major_2_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Major_2_Grade_Point']], row.iloc[columns['Major_2_Earned_Credit']],
            # cr_points
            row.iloc[columns['Major_2_Credit_Points']]
        ],
        # Row 9
        [
            # _, _, _, Major_1_Practical_subject, main, cce
            '', '', '', ('     ' + str(row.iloc[columns['Major_2_Practical_Subject']]) if pd.notna(row.iloc[columns['Major_2_Practical_Subject']]) else ''), '', '', '', #row.iloc[columns['Major_1_Practical_Main_Marks']], row.iloc[columns['Major_1_Practical_CCE_Marks']],
            # obt,
            row.iloc[columns['Major_2_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Major_2_Practical_Max_Credits']], row.iloc[columns['Major_2_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Major_2_Practical_Grade_Point']], row.iloc[columns['Major_2_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Major_2_Practical_Credit_Points']]
        ],
        # Row 10
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
        # Row 11
        [
            # _, _, _, Major_1_Practical_subject, main, cce
            '', '', '', ('     ' + str(row.iloc[columns['Minor_Practical_Subject']]) if pd.notna(row.iloc[columns['Minor_Practical_Subject']]) else ''), '', '', '', #row.iloc[columns['Major_1_Practical_Main_Marks']], row.iloc[columns['Major_1_Practical_CCE_Marks']],
            # obt,
            row.iloc[columns['Minor_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Minor_Practical_Max_Credits']], row.iloc[columns['Minor_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Minor_Practical_Grade_Point']], row.iloc[columns['Minor_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Minor_Practical_Credit_Points']]
        ],
        # Row 12
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
        # Row 13
        [
            # _, _, _, Major_1_Practical_subject, main, cce
            '', '', '', ('     ' + str(row.iloc[columns['Open_Practical_Subject']]) if pd.notna(row.iloc[columns['Open_Practical_Subject']]) else ''), '', '', '', #row.iloc[columns['Major_1_Practical_Main_Marks']], row.iloc[columns['Major_1_Practical_CCE_Marks']],
            # obt,
            row.iloc[columns['Open_Practical_Total_Marks']],
            # max_credits, letter_grade,
            row.iloc[columns['Open_Practical_Max_Credits']], row.iloc[columns['Open_Practical_Grade_Letter']],
            # grade_point, earned_creadit,
            row.iloc[columns['Open_Practical_Grade_Point']], row.iloc[columns['Open_Practical_Earned_Credit']],
            # cr_points
            row.iloc[columns['Open_Practical_Credit_Points']]
        ],
        # Row 14
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
        # Row 15
        [
            # _, _, _, VOC_practical, main,
            '', '', '', ('     ' + str(row.iloc[columns['VOC_Practical']]) if pd.notna(row.iloc[columns['VOC_Practical']]) else ''), '', '', #row.iloc[columns['VOC_Practical_Total_marks']],
            #  cce, obt,
            '', row.iloc[columns['VOC_Practical_Total_marks']],
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
        # Row 17
        [
            # _, _, _, _, 'TOTAL'
            '', '', '', '                                                               TOTAL:', '',
            # _, _, _, total max credits
            '', '', '', row.iloc[columns['Total_Max_Credits']],
            # _, Total earned creadits, total credit points
            '', '', row.iloc[columns['Total_Earned_Credits']], row.iloc[columns['Total_Credit_Points']]
        ],
        # Row 18
        [
            # _, 'AGPA = CREDIT POINTS/MAX CREDIT', _, _, 'AGPA : agpa var'
            '', 'AGPA = CREDIT POINTS/MAX CREDIT', '', '                                                                 AGPA:', 
            '' + str(row.iloc[columns['AGPA_3']]),
            # _, _, _
            '', '', '', '',
            # _, _, _, _
            '', '', ''
        ],
        # Row 19
        [
            # _, 'CATEG.: category var', _, _, 'RESULT: result var'
            'CATEG.:' + row.iloc[columns['Category']], '', '            AGPA1:    ' + str(row.iloc[columns['AGPA_1']]), '     AGPA2:    ' + str(row.iloc[columns['AGPA_2']]) + '            AGPA3:    ' + str(row.iloc[columns['AGPA_3']]),
            # _, _, _
            'CGPA:  ' + str(row.iloc[columns['CGPA']]), '', '', 'PER.:  ' + str(row.iloc[columns['Percentage']]), '', 'RESULT:' + row.iloc[columns['Result']],
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

row_idx = 1
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
