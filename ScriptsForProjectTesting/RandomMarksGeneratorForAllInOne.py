"""
Script: Random Marks Generator

Description:
------------
This script populates student data in specified Excel files with random marks for main marks, CCE and practical marks. 
The results are saved to new Excel files with a suffix indicating the marks have been updated. Any errors during processing are logged for review.

Steps:
------
1. Define a list of Excel file paths that contain student enrollment data.
3. Create a function to generate random marks for each student:
   - `main_marks`: Random value between 0 and 70, or "AB", or empty (1 in 10 chance for "AB" and 1 in 10 chance for empty).
   - `cce`: Random value between 0 and 30, or "AB", or empty (1 in 10 chance for "AB" and 1 in 10 chance for empty).
   - `practical_marks`: Random value between 0 and 100, or "AB", or empty (1 in 10 chance for "AB" and 1 in 10 chance for empty).

"""


import pandas as pd
import random
import os
import logging

# Configure logging
logging.basicConfig(
    filename='/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/random_marks_generator_all_files.log',  # Log file name
    level=logging.ERROR,                # Set the logging level to ERROR
    format='%(asctime)s:%(levelname)s:%(message)s'
)

# Step 1 : List of Excel file paths to populate the data
file_paths = [
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Geography.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/History.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Physics.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Sociology.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_1/Chemistry.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_1/Computer Science.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_1/Psychology.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_2/Biology.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_2/Economics.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_2/Political Science.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/minor/Accounting.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/minor/Banking.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/minor/English.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/minor/Marketing.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/open/Event Management.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/open/Fashion Design.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/open/Maths.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/open/Tourism.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/project/internship/Internship.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/project/internship/Project.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/voc/Graphic Design.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/voc/IT.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/voc/Photography.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/voc/Retail Management.xlsx'
]

def generate_random_marks(row):
    # Define the ranges for the marks
    main_marks_range = (0, 70)
    cce_range = (0, 30)
    practical_marks_range = (0, 100)

    # Function to generate a random mark or special entry
    def random_marks_with_special_entries(range_tuple):
        # Generate a random number
        random_number = random.randint(*range_tuple)
        
        # Use a weighted choice to determine the output
        choice = random.choices(
            [random_number, 'AB', ''],  # Choices
            weights=[8, 1, 1],          # Probabilities: 8/10 for number, 1/10 for AB, 1/10 for empty
            k=1                         # Number of choices to return
        )
        
        return choice[0]

    # Apply the function to each marks field
    row['main_marks'] = random_marks_with_special_entries(main_marks_range)
    row['cce'] = random_marks_with_special_entries(cce_range)
    row['practical_marks'] = random_marks_with_special_entries(practical_marks_range)

    return row

# Process each file
for file_path in file_paths:
    try:
        df = pd.read_excel(file_path)
        
        # Apply the random marks generation to each row
        df = df.apply(generate_random_marks, axis=1)
        
        # Construct the output file path
        base, ext = os.path.splitext(file_path)
        output_file_path = f"{base}_marksupdated{ext}"
        
        # Save the updated DataFrame to a new Excel file
        df.to_excel(output_file_path, index=False)
    
    except Exception as e:
        # Log the error message along with the file path
        logging.error(f"Failed to process {file_path}: {e}")
