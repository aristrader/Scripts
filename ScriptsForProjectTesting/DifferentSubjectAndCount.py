"""
Script: Subject Count Extractor

Description:
------------
This script calculates the count of unique subjects from different subject type columns in an Excel file.

Steps:
------
1. Define the file path of the Excel file (file_path).
2. Define the subject type columns you want to analyze (subject_types).
3. Define the path where you want to save the results text file.

"""


import pandas as pd

# Step 1 -> Update with your file path
file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_data.xlsx'
df = pd.read_excel(file_path)

# Step 2 -> How many different subject type columns are there name them here.
subject_types = ['foundation_course', 'major_1', 'major_2', 'minor', 'open', 'voc', 'project/internship']

# Function to calculate the count of unique subject names
def count_subjects_by_type(df, subject_type):
    subject_counts = df[subject_type].value_counts()
    total_count = subject_counts.sum()
    unique_count = subject_counts.nunique()
    return subject_counts, total_count, unique_count

# Step 3 : Create a text file to store the results and define the path to it.
output_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/differentsubjectandcounts.txt'
with open(output_file_path, 'w') as f:
    # Loop through each subject type and get the counts
    for subject_type in subject_types:
        subject_counts, total_count, unique_count = count_subjects_by_type(df, subject_type)
        
        # Write the subject counts and total count to the file
        f.write(f"\nCounts for {subject_type}:\n")
        f.write(subject_counts.to_string())  # Convert the counts to a string for writing
        f.write(f"\nTotal: {total_count}\n")
        f.write(f"Unique subjects: {unique_count}\n\n")

print(f"Results written to {output_file_path}")
