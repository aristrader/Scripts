# Improvements -> improve the exception handling for more detail analysis
# Error cases -> validation failure on request body,
#         duplicate roll no or enrollment no, or unknown exception occurs

# Total steps 3

import pandas as pd
import logging
import requests

# Step 1 -> Replace this with actual courseDetailsId
course_details_id = 15
max_credits_subject = 4  # Placeholder value, replace with actual max credits (This we need to remove the constraint in the application)

# Step 2: Update with your file path for the student and subject data
file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_data.xlsx'
df = pd.read_excel(file_path)

# Step 3: Define subject type columns
subject_types = ['foundation_course', 'major_1', 'major_2', 'minor', 'open', 'voc', "project/internship"]

api_url = "http://localhost:8080/v2/subject-details/add"


logging.basicConfig(filename='/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/subject_upload_failures.log',
                    level=logging.ERROR, format='%(asctime)s %(levelname)s: %(message)s')

# Function to make API call to upload subject
def upload_subject(subject_type, subject_name):
    subject_data = {
        "courseDetailsId": course_details_id,
        "subjectType": subject_type,
        "optionsName": subject_name,
        "maxCreditsSubject": max_credits_subject
    }

    try:
        response = requests.post(api_url, json=subject_data)
        if response.status_code == 201:
            return True
        else:
            logging.error(f"Failed to upload subject {subject_name} under {subject_type}: {response.text}")
            return False
    except Exception as e:
        logging.error(f"Exception while uploading subject {subject_name} under {subject_type}: {str(e)}")
        return False

# Extract unique subjects from each subject type and make API calls
successful_uploads = 0
failed_uploads = 0

for subject_type in subject_types:
    unique_subjects = df[subject_type].unique()
    
    for subject_name in unique_subjects:
        # Make the API call to upload the subject
        if upload_subject(subject_type, subject_name):
            successful_uploads += 1
        else:
            failed_uploads += 1

print(f"Total successful uploads: {successful_uploads}")
print(f"Total failed uploads: {failed_uploads}")
