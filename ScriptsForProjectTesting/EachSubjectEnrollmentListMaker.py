# Improvements -> improve the exception handling for more detail analysis
# Error cases -> validation failure on request body,
#         duplicate roll no or enrollment no, or unknown exception occurs

# Total steps 1

import os
import pandas as pd
import logging
import requests

# Set up logging
log_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/subject_enrollment_list_failure.log'
logging.basicConfig(filename=log_file_path, level=logging.ERROR,
                    format='%(asctime)s %(levelname)s: %(message)s')

# Set up logging for summary details
summary_log_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/subject_summary.log'

# Create a logger specifically for the summary log
summary_logging = logging.getLogger('summary_logger')
summary_logging.setLevel(logging.INFO)  # Set the level to INFO to capture info logs

# File handler for writing summary logs to file
summary_handler = logging.FileHandler(summary_log_file_path)
summary_handler.setLevel(logging.INFO)  # Ensure the handler is set to INFO level

# Formatter for the log file entries
summary_formatter = logging.Formatter('%(asctime)s - %(message)s')
summary_handler.setFormatter(summary_formatter)

# Add handler to the summary logger
summary_logging.addHandler(summary_handler)

# Flush the handler to ensure logs are written
summary_logging.propagate = False

# API URLs
BASE_URL = "http://localhost:8080/v2"
GET_SUBJECTS_URL = f"{BASE_URL}/subject-details/course/{{course_details_id}}"
GET_ENROLLMENTS_URL = f"{BASE_URL}/enrollment-and-marks/getBySubjectDetailsId/{{subject_details_id}}"
GET_STUDENT_URL = f"{BASE_URL}/students/get/id/{{student_id}}"

# Directory for output
output_directory = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/ListsSubjectWise'

# Step 1: Replace with actual course details ID
course_details_id = 15

# Fetch all subjects
try:
    subjects_response = requests.get(GET_SUBJECTS_URL.format(course_details_id=course_details_id))
    subjects_response.raise_for_status()
    subjects = subjects_response.json().get('data', [])
except requests.exceptions.RequestException as e:
    logging.error(f"Failed to fetch subjects for course_details_id {course_details_id}: {str(e)}")
    raise SystemExit(f"Critical error: Unable to proceed due to API failure: {str(e)}")

# Group subjects by subject type
subjects_by_type = {}
for subject in subjects:
    subject_type = subject['subjectType']
    options_name = subject['optionsName']
    subject_details_id = subject['subjectDetailsId']
    
    if subject_type not in subjects_by_type:
        subjects_by_type[subject_type] = {}
    
    subjects_by_type[subject_type][options_name] = subject_details_id

# Iterate over each subject type and each option
for subject_type, options in subjects_by_type.items():
    # Create a directory for the subject type if it doesn't exist
    subject_type_dir = os.path.join(output_directory, subject_type)
    os.makedirs(subject_type_dir, exist_ok=True)
    
    # Initialize summary data for logging
    summary_data = {}
    
    for options_name, subject_details_id in options.items():
        # Fetch all enrollments for the current subject option
        try:
            enrollments_response = requests.get(GET_ENROLLMENTS_URL.format(subject_details_id=subject_details_id))
            enrollments_response.raise_for_status()
            enrollments = enrollments_response.json().get('data', [])
        except requests.exceptions.RequestException as e:
            logging.error(f"Failed to fetch enrollments for subject_details_id {subject_details_id}: {str(e)}")
            continue  # Skip to the next subject option

        # Fetch student details for each enrollment
        students_data = []
        serial_no = 1
        for enrollment in enrollments:
            student_id = enrollment['studentId']
            
            try:
                student_response = requests.get(GET_STUDENT_URL.format(student_id=student_id))
                student_response.raise_for_status()
                student = student_response.json().get('data', {})
                
                student_info = {
                    'subject_id': subject_details_id,
                    'student_id': student_id,
                    'S.No': serial_no,
                    'roll_no': student['rollNo'],
                    'enrollment_no': student['enrollmentNo'],
                    'student_name': student['studentName'],
                    'main_marks': '',  # Empty column for main marks
                    'cce': '',         # Empty column for CCE
                    'practical_marks': ''  # Empty column for practical marks
                }
                students_data.append(student_info)
                serial_no += 1
            
            except requests.exceptions.RequestException as e:
                logging.error(f"Failed to fetch student details for student_id {student_id}: {str(e)}")
                continue  # Skip to the next student

        # Create a DataFrame and save to an Excel file
        df = pd.DataFrame(students_data)
        file_path = os.path.join(subject_type_dir, f"{options_name}.xlsx")
        
        if not df.empty:
            df.to_excel(file_path, index=False)
            print(f"Saved enrollment data for {options_name} under {subject_type} at {file_path}")
            
            summary_data[options_name] = len(students_data)
        else:
            logging.error(f"No enrollments found for subject option: {options_name} in subject type: {subject_type}")
            
    total_students = sum(summary_data.values())
    total_options = len(summary_data)
    
    summary_logging.info(f"Subject Type: {subject_type}")
    for option_name, student_count in summary_data.items():
        summary_logging.info(f"Option Name: {option_name}, Students: {student_count}")

    summary_logging.info(f"Total Option Types: {total_options}, Total Students: {total_students}\n")
    
    summary_handler.flush()

print("Processing complete.")
