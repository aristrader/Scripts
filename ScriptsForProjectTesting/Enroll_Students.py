"""
Script: Student Subject Enrollment

Description:
------------
This script automates enrolling students into their respective subjects based on data from an Excel file.

Steps:
------
1. Update the file path for the student data (file_path).
2. Ensure the correct courseDetailsId are set.
3. Define the subject type columns in the Excel file.

"""


import pandas as pd
import logging
import requests  # Assuming API calls for fetching students and subjects and making enrollments

# Step 1: Update with your file path
file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_data.xlsx'
df = pd.read_excel(file_path)

log_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/enrollment_failure.log'
logging.basicConfig(filename=log_file_path, level=logging.ERROR, 
                    format='%(asctime)s %(levelname)s: %(message)s')

# API URLs for fetching students, subjects, and enrolling students in subjects
BASE_URL = "http://localhost:8080/v2"
GET_STUDENTS_URL = f"{BASE_URL}/students/getByCourseDetailsId/{{course_details_id}}"
GET_SUBJECTS_URL = f"{BASE_URL}/subject-details/course/{{course_details_id}}"
ENROLL_URL = f"{BASE_URL}/enrollment-and-marks/add"

# Step 2: Replace with actual course details ID
course_details_id = 1

# Fetch all students and subjects by course_details_id
try:
    students_response = requests.get(GET_STUDENTS_URL.format(course_details_id=course_details_id))
    students_response.raise_for_status()  # Raises an HTTPError for bad responses (4xx/5xx)
    students = students_response.json().get('data', [])

    subjects_response = requests.get(GET_SUBJECTS_URL.format(course_details_id=course_details_id))
    subjects_response.raise_for_status()  # Raises an HTTPError for bad responses (4xx/5xx)
    subjects = subjects_response.json().get('data', [])

except requests.exceptions.RequestException as e:
    logging.error(f"Failed to fetch students or subjects for course_details_id {course_details_id}: {str(e)}")
    raise SystemExit(f"Critical error: Unable to proceed with enrollment due to API failure: {str(e)}")

if isinstance(students, list):
    students_lookup = {str(student['rollNo']).strip(): student['studentId'] for student in students}  # Using roll number as key
    students_lookup.update({str(student['enrollmentNo']).strip(): student['studentId'] for student in students})
else:
    raise Exception(f"Unexpected students data format: {type(students)}")

if isinstance(subjects, list):
    subjects_lookup = {
        (str(subject['optionsName']).strip().lower(), str(subject['subjectType']).strip().lower()): subject['subjectDetailsId']
        for subject in subjects
    }
else:
    raise Exception(f"Unexpected subjects data format: {type(subjects)}")

# Iterate over the Excel rows and process each student
total_students = len(df)
total_subject_enrollments = 0  # Track the total number of subject enrollments attempted
successful_enrollments = 0  # Track the number of successful subject enrollments
failed_enrollments = 0  # Track the number of failed subject enrollments

for index, row in df.iterrows():
    try:
        # Fetch student details by roll number or enrollment number
        roll_no = str(row['roll no']).strip()  # Ensure it's a string and remove extra spaces
        enrollment_no = str(row['enrollment no']).strip() # Ensure it's a string and remove extra spaces
        student_id = students_lookup.get(roll_no) or students_lookup.get(enrollment_no)

        if not student_id:
            raise ValueError(f"Student not found for roll no: {roll_no} or enrollment no: {enrollment_no}")

        # Step 3 : Set the appropriate subject type from the column names in excel
        subject_types = ['foundation_course', 'major_1', 'major_2', 'minor', 'open', 'voc', "project/internship"]
        
        for subject_type in subject_types:
            try:
                subject_name = str(row[subject_type]).strip()  # Ensure subject name is a string and trim spaces
                subject_key = (subject_name.lower(), subject_type.lower())  # Normalize the key for lookup
                subject_id = subjects_lookup.get(subject_key)

                if not subject_id:
                    raise ValueError(f"Subject not found for name: {subject_name} in subject type: {subject_type}")

                # Prepare the enrollment request payload
                enrollment_payload = {
                    "studentId": student_id,
                    "subjectDetailsId": subject_id,
                    "mainMarks": None,  # Add logic to fill marks if available
                    "cce": None,        # Add logic to fill CCE if available
                    "practicalMarks": None  # Add logic to fill practical marks if available
                }

                # Make the enrollment API call
                enroll_response = requests.post(ENROLL_URL, json=enrollment_payload)

                if enroll_response.status_code != 201:
                    raise ValueError(f"Failed to enroll student {roll_no} in subject {subject_name}")
                
                successful_enrollments += 1

            except Exception as subject_error:
                # Log the failure for the individual subject but continue processing the rest
                logging.error(f"Failed to enroll student {roll_no} for subject {subject_name} (type: {subject_type}): {str(subject_error)}")
                failed_enrollments += 1
                continue
            
            finally:
                # Whether successful or failed, it's an attempted enrollment, so count it
                total_subject_enrollments += 1

    except Exception as e:
        # Log the failure details into the log file
        logging.error(f"Failed to enroll student for row {index + 1}: {str(e)}")
        failed_enrollments += 1

# Print the summary of enrollments
print(f"Enrollment process completed.")
print(f"Total subject enrollments attempted: {total_subject_enrollments}")
print(f"Successful enrollments: {successful_enrollments}")
print(f"Failed enrollments: {failed_enrollments}")
