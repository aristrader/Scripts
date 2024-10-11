import pandas as pd
import requests
import logging

# Set up logging
log_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/enrollment_update_marks.log'
logging.basicConfig(filename=log_file_path, level=logging.INFO, 
                    format='%(asctime)s %(levelname)s: %(message)s')

# Step 1: Load the Excel file with populated marks
file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise/foundation_course/Geography_populated.xlsx'
df = pd.read_excel(file_path)

# API URL for updating marks
UPDATE_MARKS_URL = "http://localhost:8080/v2/enrollment-and-marks/addMarks"

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
        response = requests.post(UPDATE_MARKS_URL, json=payload)  # Using PUT or POST based on your API
        response.raise_for_status()  # Raise an error for bad responses (4xx/5xx)
        logging.info(f"Successfully updated marks for student ID {student_id}, subject ID {subject_id}")

    except requests.exceptions.RequestException as e:
        logging.error(f"Failed to update marks for student ID {student_id}, subject ID {subject_id}: {str(e)}")
