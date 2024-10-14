'''
Script: Enrollment Check Marks List

Description:
------------
This script reads student marks from specified Excel files and verifies them by fetching corresponding data from an API. 
It logs any discrepancies between the marks in the Excel files and those retrieved from the API.

Steps:
------
1. Define a list of Excel file paths that contain student marks to be checked.

'''

import pandas as pd
import requests
import logging
import os
import time

# Set up logging
log_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/enrollment_check_marks_list.log'
logging.basicConfig(filename=log_file_path, level=logging.INFO, 
                    format='%(asctime)s %(levelname)s: %(message)s')

# List of Excel file paths with populated marks
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

# API URL for fetching enrollment marks
GET_MARKS_URL = "http://localhost:8080/v2/enrollment-and-marks/getByStudentAndSubjectId/{studentId}/{subjectDetailsId}"

# Process each file
for file_path in file_paths:
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Iterate over each row and check marks
        for index, row in df.iterrows():
            student_id = row['student_id']
            subject_id = row['subject_id']

            # Prepare the URL with student and subject IDs
            url = GET_MARKS_URL.format(studentId=student_id, subjectDetailsId=subject_id)

            try:
                # Call the endpoint to fetch marks
                response = requests.get(url)  # Using GET based on the new API
                response.raise_for_status()  # Raise an error for bad responses (4xx/5xx)

                # Parse the response
                fetched_data = response.json().get('data')
                if fetched_data:
                    fetched_main_marks = fetched_data.get('mainMarks')
                    fetched_cce = fetched_data.get('cce')
                    fetched_practical_marks = fetched_data.get('practicalMarks')

                    # Handle null values from the API response
                    fetched_main_marks = fetched_main_marks if fetched_main_marks is not None else None
                    fetched_cce = fetched_cce if fetched_cce is not None else None
                    fetched_practical_marks = fetched_practical_marks if fetched_practical_marks is not None else None

                    # Handle empty entries in the Excel sheet
                    excel_main_marks = row['main_marks'] if pd.notna(row['main_marks']) else None
                    excel_cce = row['cce'] if pd.notna(row['cce']) else None
                    excel_practical_marks = row['practical_marks'] if pd.notna(row['practical_marks']) else None
                    
                    # Convert Excel marks to strings
                    excel_main_marks = str(excel_main_marks) if excel_main_marks is not None else None
                    excel_cce = str(excel_cce) if excel_cce is not None else None
                    excel_practical_marks = str(excel_practical_marks) if excel_practical_marks is not None else None

                    # Compare with the Excel values
                    if (fetched_main_marks != excel_main_marks or 
                        fetched_cce != excel_cce or 
                        fetched_practical_marks != excel_practical_marks):
                        logging.error(f"Mismatch for student ID {student_id}, subject ID {subject_id}: "
                                        f"Excel (main: {excel_main_marks}, cce: {excel_cce}, practical: {excel_practical_marks}) "
                                        f"vs API (main: {fetched_main_marks}, cce: {fetched_cce}, practical: {fetched_practical_marks})")
                    else :
                        logging.info("data match for student ID {student_id}, subject ID {subject_id} ")

                else:
                    logging.error(f"No data found for student ID {student_id}, subject ID {subject_id}")

            except requests.exceptions.RequestException as e:
                logging.error(f"Failed to fetch marks for student ID {student_id}, subject ID {subject_id}: {str(e)}")

    except Exception as e:
        logging.error(f"Failed to process file {file_path}: {str(e)}")
