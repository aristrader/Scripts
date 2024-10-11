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

# Populate random marks for each student
def generate_random_marks(row):
    # Generate random marks for main_marks, cce, and practical_marks
    row['main_marks'] = random.randint(0, 70)  # Random value between 0 and 70
    row['cce'] = random.randint(0, 30)         # Random value between 0 and 30
    row['practical_marks'] = random.randint(0, 100)  # Random value between 0 and 100
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
