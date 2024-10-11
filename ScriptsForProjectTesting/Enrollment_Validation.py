# Improvements -> improve the exception handling for more detail analysis
# Error cases -> validation failure on request body,
#         duplicate roll no or enrollment no, or unknown exception occurs

# When the data already exists then that is not an error in the enrollment script. Remove it. 
# (Scenario -> out of 100 enrollment 1 fail, you can run all 100 and you will get 99 error which is a clutter)

# Total steps 4

import pandas as pd
import requests

# API URLs
BASE_URL = "http://localhost:8080/v2"
GET_STUDENTS_URL = f"{BASE_URL}/students/getByCourseDetailsId/{{course_details_id}}"
GET_SUBJECTS_URL = f"{BASE_URL}/subject-details/course/{{course_details_id}}"
GET_ENROLLMENTS_URL = f"{BASE_URL}/enrollment-and-marks/getByStudentId/{{student_id}}"

# Step 1: Replace with actual course details ID
course_details_id = 15

# Fetch students
students_response = requests.get(GET_STUDENTS_URL.format(course_details_id=course_details_id))
students_response.raise_for_status()
students = students_response.json().get('data', [])

# Fetch subjects
subjects_response = requests.get(GET_SUBJECTS_URL.format(course_details_id=course_details_id))
subjects_response.raise_for_status()
subjects = subjects_response.json().get('data', [])

# Create a subjects lookup dictionary to map subject details ID to subject type and options name
subjects_lookup = {
    subject['subjectDetailsId']: {
        'type': subject['subjectType'],
        'options_name': subject['optionsName']
    }
    for subject in subjects
}

# Prepare a list to hold all the data for the Excel
data_for_excel = []

# Iterate through each student to fetch their enrollments
for student in students:
    student_id = student['studentId']
    
    # Fetch enrollments for the current student
    enrollments_response = requests.get(GET_ENROLLMENTS_URL.format(student_id=student_id))
    enrollments_response.raise_for_status()
    enrollments = enrollments_response.json().get('data', [])
    
    # Prepare the student info
    student_info = {
        'roll no': student['rollNo'],
        'enrollment no': student['enrollmentNo'],
        'student name': student['studentName'],
        'father name': student['fatherName'],
        'mother name': student['motherName'],
    }
    
    # STEP 2 : Initialize subject type columns (subject type names will change)
    student_info.update({
        'foundation_course': None,
        'major_1': None,
        'major_2': None,
        'minor': None,
        'open': None,
        'voc': None,
        'project/internship': None
    })

    # Populate the subject type columns based on enrollments
    for enrollment in enrollments:
        subject_id = enrollment['subjectDetailsId']
        subject_info = subjects_lookup.get(subject_id)

        if subject_info:
            subject_type = subject_info['type']
            options_name = subject_info['options_name']
            
            # Step 3 : Map the options name to the correct subject type column (Subject type will change (case sensative))
            if subject_type == 'foundation_course':
                student_info['foundation_course'] = options_name
            elif subject_type == 'major_1':
                student_info['major_1'] = options_name
            elif subject_type == 'major_2':
                student_info['major_2'] = options_name
            elif subject_type == 'minor':
                student_info['minor'] = options_name
            elif subject_type == 'open':
                student_info['open'] = options_name
            elif subject_type == 'voc':
                student_info['voc'] = options_name
            elif subject_type == 'project/internship':
                student_info['project/internship'] = options_name
            # Add any additional subject types as needed

    # Append the student info to the data list
    data_for_excel.append(student_info)

# Create a DataFrame and save it to Excel
df = pd.DataFrame(data_for_excel)

# Step 4 : Update the excel path
output_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_verification.xlsx'
df.to_excel(output_file_path, index=False)

print(f"Verification report created successfully at {output_file_path}")
