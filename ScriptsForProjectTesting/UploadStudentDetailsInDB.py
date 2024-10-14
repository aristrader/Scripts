"""
Script: Student Data Uploader

Description:
------------
This script reads student data from an Excel file and uploads it to the database

Steps:
------
1. Replace with the actual courseDetailsId (courseDetailId).
2. Define the file path for the student data Excel file (file_path).
3. Ensure the column names in the Excel file match the expected format in the script.

"""

import pandas as pd
import requests
import json
import logging

# Step 1 -> Replace this with actual courseDetailsId
courseDetailId = 1

# Step 2 -> Update with your file path for the student data
file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_data.xlsx'
api_url = 'http://localhost:8080/v2/students/add'

df = pd.read_excel(file_path)

# Logging file
logging.basicConfig(filename='/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_details_upload_failures.log', 
                    level=logging.ERROR, format='%(asctime)s %(levelname)s: %(message)s')

# Initialize counters
success_count = 0
failure_count = 0

# Function to send student data to API
def save_student(student_data):
    global success_count, failure_count
    try:
        response = requests.post(api_url, json=student_data)
        if response.status_code == 201:  # 201 is the status code in a success response for this api.
            success_count += 1
        else:
            failure_count += 1
            log_failure(student_data, response.text)
    except Exception as e:
        failure_count += 1
        log_failure(student_data, str(e))

# Logging of failed student upload details
def log_failure(student_data, error_message):
    logging.error(f"Failed to save student: {student_data}, Error: {error_message}")

# Step 3 -> Make sure the column names are right in this.
for index, row in df.iterrows():
    student_data = {
        "studentName": row["student name"],
        "fatherName": row["father name"],
        "motherName": row["mother name"],
        "rollNo": str(row["roll no"]),
        "enrollmentNo": str(row["enrollment no"]),
        "courseDetailsId": courseDetailId,
        "photo": None  # Assuming photo is optional, set to None if not provided
    }
    
    # Call the function to save student
    save_student(student_data)

# Print final summary
print(f"Number of students successfully saved: {success_count}")
print(f"Number of students failed to save: {failure_count}")
