# Total steps - 2

import pandas as pd
import random

# Step 1: Give the excel path to populate the data
file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Geography.xlsx'
df = pd.read_excel(file_path)

# Populate random marks for each student
def generate_random_marks(row):
    # Generate random marks for main_marks, cce, and practical_marks
    row['main_marks'] = random.randint(0, 70)  # Random value between 0 and 70
    row['cce'] = random.randint(0, 30)         # Random value between 0 and 30
    
    # Remove this below for the test case of where practical marks are not there.
    row['practical_marks'] = random.randint(0, 100)  # Random value between 0 and 100
    return row

# Apply the random marks generation to each row
df = df.apply(generate_random_marks, axis=1)

# Step 2: Give the path to where we want to save excel and with what name
output_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Geography_populated.xlsx'  # Replace with the desired output path
df.to_excel(output_file_path, index=False)

print(f"Updated Excel file saved to: {output_file_path}")
