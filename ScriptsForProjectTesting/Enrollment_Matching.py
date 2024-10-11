# Improvements -> 

# Total steps 4

import pandas as pd

# Step 1 : define the excel paths
sheet1_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_data.xlsx'
sheet2_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_verification.xlsx'

sheet1 = pd.read_excel(sheet1_path)
sheet2 = pd.read_excel(sheet2_path)

# Drop S.no column from sheet1 if it exists
if 's.no' in sheet1.columns:
    sheet1 = sheet1.drop(columns=['s.no'])

# # Convert columns to lower case for easier comparison - Don't lower case for now.
# sheet1.columns = sheet1.columns.str.lower().str.strip()
# sheet2.columns = sheet2.columns.str.lower().str.strip()

# Step 2 : define the column on basis of which you will search and match in both sheets
sheet1.set_index('roll no', inplace=True)
sheet2.set_index('roll no', inplace=True)

# Step 3 : define the columns
columns_to_compare = ['enrollment no', 'student name', 'father name', 'mother name', 
                      'foundation_course', 'major_1', 'major_2', 'minor', 'open', 'voc', 'project/internship']

# Logs to store mismatch or missing entries
missing_in_sheet2 = []
mismatched_entries = []
extra_in_sheet2 = []

# Check for rows in Sheet 1 that are missing in Sheet 2
for roll_no, row1 in sheet1.iterrows():
    if roll_no not in sheet2.index:
        missing_in_sheet2.append(f"Row with roll no {roll_no} is missing in Sheet 2.")
    else:
        # Step 2: Compare each column for mismatches
        row2 = sheet2.loc[roll_no]
        for column in columns_to_compare:
            if pd.isna(row1[column]) and pd.isna(row2[column]):
                continue  # Skip comparison if both are NaN
            elif row1[column] != row2[column]:
                mismatched_entries.append(
                    f"Mismatch for roll no {roll_no} in column '{column}': Sheet 1 = '{row1[column]}', Sheet 2 = '{row2[column]}'"
                )

# Check for extra entries in Sheet 2 that are not in Sheet 1
for roll_no in sheet2.index:
    if roll_no not in sheet1.index:
        extra_in_sheet2.append(f"Row with roll no {roll_no} is extra in Sheet 2.")

# Step 4: Log the results
with open('/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/mismatch_log.txt', 'w') as log_file:
    log_file.write("Missing rows in Sheet 2:\n")
    if missing_in_sheet2:
        for entry in missing_in_sheet2:
            log_file.write(entry + '\n')
    else:
        log_file.write("None\n")
    
    log_file.write("\nMismatched entries:\n")
    if mismatched_entries:
        for entry in mismatched_entries:
            log_file.write(entry + '\n')
    else:
        log_file.write("None\n")
    
    log_file.write("\nExtra rows in Sheet 2:\n")
    if extra_in_sheet2:
        for entry in extra_in_sheet2:
            log_file.write(entry + '\n')
    else:
        log_file.write("None\n")

print("Data comparison completed. Check the log file for details.")
