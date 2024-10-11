import pandas as pd
import requests
import logging
import os

# Set up logging
log_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/enrollment_update_marks_list.log'
logging.basicConfig(filename=log_file_path, level=logging.INFO, 
                    format='%(asctime)s %(levelname)s: %(message)s')

# Step 1: List of Excel file paths with populated marks
file_paths = [
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Geography_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/History_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Physics_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Sociology_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_1/Chemistry_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_1/Computer Science_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_1/Psychology_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_2/Biology_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_2/Economics_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/major_2/Political Science_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/minor/Accounting_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/minor/Banking_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/minor/English_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/minor/Marketing_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/open/Event Management_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/open/Fashion Design_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/open/Maths_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/open/Tourism_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/project/internship/Internship_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/project/internship/Project_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/voc/Graphic Design_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/voc/IT_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/voc/Photography_marksupdated.xlsx',
    '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/voc/Retail Management_marksupdated.xlsx'
]

# API URL for updating marks
UPDATE_MARKS_URL = "http://localhost:8080/v2/enrollment-and-marks/addMarks"

# Process each file
for file_path in file_paths:
    try:
        # Load the Excel file with populated marks
        df = pd.read_excel(file_path)

        # Iterate over each row and update marks
        for index, row in df.iterrows():
            student_id = row['student_id']
            subject_id = row['subject_id']
            main_marks = row['main_marks'] if pd.notna(row['main_marks']) else None  # Set to None if empty
            cce = row['cce'] if pd.notna(row['cce']) else None  # Set to None if empty
            practical_marks = row['practical_marks'] if pd.notna(row['practical_marks']) else None  # Set to None if empty

            # Prepare the payload for the API call
            payload = {
                "studentId": student_id,
                "subjectDetailsId": subject_id,
                "mainMarks": main_marks,
                "cce": cce,
                "practicalMarks": practical_marks
            }

            try:
                # Call the endpoint to update marks
                response = requests.post(UPDATE_MARKS_URL, json=payload)  # Using POST based on your API
                response.raise_for_status()  # Raise an error for bad responses (4xx/5xx)

            except requests.exceptions.RequestException as e:
                logging.error(f"Failed to update marks for student ID {student_id}, subject ID {subject_id}: {str(e)}")

    except Exception as e:
        logging.error(f"Failed to process file {file_path}: {str(e)}")
