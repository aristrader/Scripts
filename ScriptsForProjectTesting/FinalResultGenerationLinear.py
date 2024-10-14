import requests
import logging
import openpyxl
from openpyxl.styles import Alignment
import time

# Step 1 : replace with the actual course detail id
course_details_id = 1

# Set up logging
log_file_path = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/LOGS_FINAL_RESULTS.log'
logging.basicConfig(filename=log_file_path, level=logging.INFO, 
                    format='%(asctime)s %(levelname)s: %(message)s')

# Base URL for the college results application
BASE_URL = "http://localhost:8080/v2"

def fetch_students_enrolled(course_details_id):
    """Fetch the list of students enrolled in a given course."""
    try:
        response = requests.get(f"{BASE_URL}/students/getByCourseDetailsId/{course_details_id}")
        response.raise_for_status()
        return response.json().get('data', [])
    except Exception as e:
        logging.error(f"Failed to fetch students for course ID {course_details_id}: {str(e)}")
        return []

def fetch_enrollments_for_student(student_id):
    """Fetch the list of enrollments for a given student."""
    try:
        response = requests.get(f"{BASE_URL}/enrollment-and-marks/getByStudentId/{student_id}")
        response.raise_for_status()
        return response.json().get('data', [])
    except Exception as e:
        logging.error(f"Failed to fetch enrollments for student ID {student_id}: {str(e)}")
        return []
    
def fetch_subject_details(subject_details_id):
    """Fetch the details of a subject by its ID."""
    try:
        response = requests.get(f"{BASE_URL}/subject-details/get/id/{subject_details_id}")
        response.raise_for_status()
        return response.json().get('data', {})
    except Exception as e:
        logging.error(f"Failed to fetch subject details for subject ID {subject_details_id}: {str(e)}")
        return {}


# NOTE : THE BELOW LOGIC WILL FAIL WHEN THE NUMBERS CONTAIN A '.' IN THEM. IN THAT CASE YOU NEED CUSTOM FUNCTIONS
def calculate_total_marks(main_marks, cce_marks):
    """Calculate total marks based on conditions for main and CCE marks."""
    # If both are None, return None
    if main_marks is None and cce_marks is None:
        return None
    
    # If both are alphabetic strings, return main_marks
    if isinstance(main_marks, str) and main_marks.isalpha() and isinstance(cce_marks, str) and cce_marks.isalpha():
        return main_marks

    # If one is an alphabetic string and the other is None, return the alphabetic string
    if isinstance(main_marks, str) and main_marks.isalpha() and cce_marks is None:
        return main_marks
    if isinstance(cce_marks, str) and cce_marks.isalpha() and main_marks is None:
        return cce_marks

    # If one is numeric string and the other is None, return the numeric string
    if main_marks is None and isinstance(cce_marks, str) and cce_marks.isdigit():
        return cce_marks
    if cce_marks is None and isinstance(main_marks, str) and main_marks.isdigit():
        return main_marks

    # If one is numeric string and the other is alphabetic string, return the numeric string
    if isinstance(main_marks, str) and main_marks.isdigit() and isinstance(cce_marks, str) and cce_marks.isalpha():
        return main_marks
    if isinstance(cce_marks, str) and cce_marks.isdigit() and isinstance(main_marks, str) and main_marks.isalpha():
        return cce_marks

    # If both are numeric strings, return the string sum of their numeric values
    if isinstance(main_marks, str) and main_marks.isdigit() and \
       isinstance(cce_marks, str) and cce_marks.isdigit():
        return str(int(main_marks) + int(cce_marks))

    return None  # Default return for any unhandled case


def calculate_grade_and_grade_points(total_marks):
    """Determine the grade based on the conditions for total marks."""
    # If total is None, return None and grade point as 0
    if total_marks is None:
        return "F", 0

    # If total is an alphabetic string, return the same string and grade point as 0
    if isinstance(total_marks, str) and total_marks.isalpha():
        return total_marks, 0

    # If total is a numeric string, apply grading logic
    if isinstance(total_marks, str) and total_marks.isdigit():
        total_marks = int(total_marks)  # Convert to integer for grading
        if total_marks >= 90:
            return "A+", 9  # Highest grade
        elif total_marks >= 80:
            return "A", 8
        elif total_marks >= 70:
            return "B+", 7
        elif total_marks >= 60:
            return "B", 6
        elif total_marks >= 50:
            return "C", 5
        elif total_marks >= 40:
            return "P", 4
        else:
            return "F", 0  # Fail grade

    return None  # Default return for any unhandled case

def calculate_earned_credits(grade, max_credits):
    """Determine earned credits based on the grade."""
    # Check if the grade is one of the accepted grades (A+, A, B+, B, C, P)
    if grade in ['A+', 'A', 'B+', 'B', 'C', 'P']:
        return max_credits if max_credits is not None else 0  # Return max credits for valid grades
    else:
        return 0  # Return 0 for all other cases including F, None, or invalid strings


def calculate_credit_points(earned_credits, grade_points):
    """Calculate the credit points as earned credits * grade points."""
    return earned_credits * grade_points

def determine_result(subject_grades):
    """Determine the final result based on the number of failed subjects."""
    failed_subjects = sum(1 for grade, _ in subject_grades if grade == "F")
    
    if failed_subjects > 2:
        return "Fail"
    elif failed_subjects > 0:
        return "Supply"
    else:
        return "Pass"

def process_course_results(course_details_id):
    """Process the student results for a course."""
    students = fetch_students_enrolled(course_details_id)
    final_results = []
    
    # Process each student
    try:
        sorted_students = sorted(students, key=lambda x: int(x['rollNo']))  # Sorting by roll number
    except ValueError as e:
        logging.error(f"Error sorting students: {e}")
        return  # Exit the function if sorting fails


    # Process each student
    for student in sorted_students:
        enrollments = fetch_enrollments_for_student(student['studentId'])
        
        total_max_credits = 0
        total_earned_credits = 0
        total_credit_points = 0
        subject_grades = []

        # Create a dictionary to hold subject results
        subject_results = {
            "foundation_course": {},
            "major_1": {},
            "major_2": {},
            "minor": {},
            "open": {},
            "voc": {},
            "project/internship": {},
            "foundation_course_practical": {},
            "major_1_practical": {},
            "major_2_practical": {},
            "minor_practical": {},
            "open_practical": {},
            "voc_practical": {},
            "project/internship_practical": {}
        }

        # Process each enrollment (subject)
        for enrollment in enrollments:
            subject = fetch_subject_details(enrollment['subjectDetailsId'])  # Get subject info
            
            main_marks = enrollment['mainMarks']
            cce_marks = enrollment['cce']
            total_marks = calculate_total_marks(main_marks, cce_marks)
            
            max_credits = subject['maxCreditsSubject']
            grade, grade_points = calculate_grade_and_grade_points(total_marks)
            earned_credits = calculate_earned_credits(grade, max_credits)
            credit_points = calculate_credit_points(earned_credits, grade_points)

            # Update totals
            total_max_credits += max_credits if max_credits is not None else 0
            total_earned_credits += earned_credits
            total_credit_points += credit_points
            subject_grades.append((grade, grade_points))

            # Store subject results based on subject type
            subject_results[subject['subjectType']] = {
                "subject_name": subject['optionsName'],
                "main_marks": main_marks,
                "cce_marks": cce_marks,
                "total_marks": total_marks,
                "max_credits": max_credits,
                "grade": grade,
                "grade_points": grade_points,
                "earned_credits": earned_credits,
                "credit_points": credit_points
            }

            # Check for practical marks and store separately
            if 'practicalMarks' in enrollment and enrollment['practicalMarks'] is not None:
                practical_marks = enrollment['practicalMarks']
                practical_total_marks = calculate_total_marks(practical_marks, None)  # Assuming CCE is not applicable here
                
                # Store practical results
                subject_results[subject['subjectType'] + '_practical'] = {
                    "subject_name": "Practical",
                    "main_marks": practical_marks,
                    "cce_marks": None,  # No CCE for practical
                    "total_marks": practical_total_marks,
                    "max_credits": max_credits,  # Assuming same max credits
                    "grade": calculate_grade_and_grade_points(practical_total_marks)[0],  # Get only grade
                    "grade_points": calculate_grade_and_grade_points(practical_total_marks)[1],  # Get only grade points
                    "earned_credits": calculate_earned_credits(calculate_grade_and_grade_points(practical_total_marks)[0], max_credits),
                    "credit_points": calculate_credit_points(earned_credits, grade_points)
                }
            else:
                # Fill practical subject results with empty values
                subject_results[subject['subjectType'] + '_practical'] = {
                    "subject_name": None,
                    "main_marks": None,
                    "cce_marks": None,
                    "total_marks": None,
                    "max_credits": None,
                    "grade": None,
                    "grade_points": None,
                    "earned_credits": None,
                    "credit_points": None
                }

        # Final result (pass/fail/supply)
        result_status = determine_result(subject_grades)

        # Add the student data to the final results list in linear format
        final_results.append({
            "roll_no": student['rollNo'],
            "enrollment_no": student['enrollmentNo'],
            "student_name": student['studentName'],
            "father_name": student['fatherName'],
            "mother_name": student['motherName'],
            "foundation_course": subject_results["foundation_course"],
            "major_1": subject_results["major_1"],
            "major_2": subject_results["major_2"],
            "minor": subject_results["minor"],
            "open": subject_results["open"],
            "voc": subject_results["voc"],
            "project/internship": subject_results["project/internship"],
            "foundation_course_practical": subject_results["foundation_course_practical"],
            "major_1_practical": subject_results["major_1_practical"],
            "major_2_practical": subject_results["major_2_practical"],
            "minor_practical": subject_results["minor_practical"],
            "open_practical": subject_results["open_practical"],
            "voc_practical": subject_results["voc_practical"],
            "project/internship_practical": subject_results["project/internship_practical"],
            "total_max_credits": total_max_credits,
            "total_earned_credits": total_earned_credits,
            "total_credit_points": total_credit_points,
            "result": result_status,
            "category": "Regular"
        })

    return final_results

def write_to_excel(final_results, file_name):
    """Write the processed results to an Excel file in the required format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Course Results"
    
    # Set headers
    headers = [
        "Roll No", "Enrollment No", "Student Name", "Father Name", "Mother Name",
        # Foundation
        "Foundation Subject Name", "Foundation Main Marks", "Foundation CCE Marks", "Foundation Total Marks", 
        "Max Credits Foundation", "Letter Grade Foundation", "Grade Point for Foundation", 
        "Earned Credit for Foundation", "Credit Points (Earned Credit * Grade Point) for Foundation",
        "Foundation Practical", "Foundation Practical Main Marks", "Foundation Practical CCE Marks", "Foundation Practical Total Marks", 
        "Max Credits Foundation Practical", "Letter Grade Foundation Practical", "Grade Point for Foundation Practical", 
        "Earned Credit for Foundation Practical", "Credit Points (Earned Credit * Grade Point) for Foundation Practical",
        # Major 1
        "Major_1 Subject Name", "Major_1 Main Marks", "Major_1 CCE Marks", "Major_1 Total Marks", 
        "Max Credits Major_1", "Letter Grade Major_1", "Grade Point for Major_1", 
        "Earned Credit for Major_1", "Credit Points (Earned Credit * Grade Point) for Major_1",
        "Major_1 Practical", "Major_1 Practical Main Marks", "Major_1 Practical CCE Marks", "Major_1 Practical Total Marks", 
        "Max Credits Major_1 Practical", "Letter Grade Major_1 Practical", "Grade Point for Major_1 Practical", 
        "Earned Credit for Major_1 Practical", "Credit Points (Earned Credit * Grade Point) for Major_1 Practical",
        # Major 2
        "Major_2 Subject Name", "Major_2 Main Marks", "Major_2 CCE Marks", "Major_2 Total Marks", 
        "Max Credits Major_2", "Letter Grade Major_2", "Grade Point for Major_2", 
        "Earned Credit for Major_2", "Credit Points (Earned Credit * Grade Point) for Major_2",
        "Major_2 Practical", "Major_2 Practical Main Marks", "Major_2 Practical CCE Marks", "Major_2 Practical Total Marks", 
        "Max Credits Major_2 Practical", "Letter Grade Major_2 Practical", "Grade Point for Major_2 Practical", 
        "Earned Credit for Major_2 Practical", "Credit Points (Earned Credit * Grade Point) for Major_2 Practical",
        # Minor
        "Minor Subject Name", "Minor Main Marks", "Minor CCE Marks", "Minor Total Marks", 
        "Max Credits Minor", "Letter Grade Minor", "Grade Point for Minor", 
        "Earned Credit for Minor", "Credit Points (Earned Credit * Grade Point) for Minor",
        "Minor Practical", "Minor Practical Main Marks", "Minor Practical CCE Marks", "Minor Practical Total Marks", 
        "Max Credits Minor Practical", "Letter Grade Minor Practical", "Grade Point for Minor Practical", 
        "Earned Credit for Minor Practical", "Credit Points (Earned Credit * Grade Point) for Minor Practical",
        # Open
        "Open Subject Name", "Open Main Marks", "Open CCE Marks", "Open Total Marks", 
        "Max Credits Open", "Letter Grade Open", "Grade Point for Open", 
        "Earned Credit for Open", "Credit Points (Earned Credit * Grade Point) for Open",
        "Open Practical", "Open Practical Main Marks", "Open Practical CCE Marks", "Open Practical Total Marks", 
        "Max Credits Open Practical", "Letter Grade Open Practical", "Grade Point for Open Practical", 
        "Earned Credit for Open Practical", "Credit Points (Earned Credit * Grade Point) for Open Practical",
        # Vocational
        "VOC Subject Name", "VOC Main Marks", "VOC CCE Marks", "VOC Total Marks", 
        "Max Credits VOC", "Letter Grade VOC", "Grade Point for VOC", 
        "Earned Credit for VOC", "Credit Points (Earned Credit * Grade Point) for VOC",
        "VOC Practical", "VOC Practical Main Marks", "VOC Practical CCE Marks", "VOC Practical Total Marks", 
        "Max Credits VOC Practical", "Letter Grade VOC Practical", "Grade Point for VOC Practical", 
        "Earned Credit for VOC Practical", "Credit Points (Earned Credit * Grade Point) for VOC Practical",
        # Project
        "Project Subject Name","Project Main Marks", "Project CCE Marks", "Project Total Marks", 
        "Max Credits Project", "Letter Grade Project", "Grade Point for Project", 
        "Earned Credit for Project", "Credit Points (Earned Credit * Grade Point) for Project",
        "Project Practical", "Project Practical Main Marks", "Project Practical CCE Marks", "Project Practical Total Marks", 
        "Max Credits Project Practical", "Letter Grade Project Practical", "Grade Point for Project Practical", 
        "Earned Credit for Project Practical", "Credit Points (Earned Credit * Grade Point) for Project Practical",
        # Totals
        "Total of Max Credits", "Total Credit Obtained", "Total Credit Points Obtained", 
        "Result", "Category"
    ]
    
    ws.append(headers)

    # Fill data
    for result in final_results:
        
        row = [
            result['roll_no'],
            result['enrollment_no'],
            result['student_name'],
            result['father_name'],
            result['mother_name'],
            # Foundation
            result.get('foundation_course', {}).get('subject_name', ''),
            result.get('foundation_course', {}).get('main_marks', ''),
            result.get('foundation_course', {}).get('cce_marks', ''),
            result.get('foundation_course', {}).get('total_marks', ''),
            result.get('foundation_course', {}).get('max_credits', ''),
            result.get('foundation_course', {}).get('grade', ''),
            result.get('foundation_course', {}).get('grade_points', ''),
            result.get('foundation_course', {}).get('earned_credits', ''),
            result.get('foundation_course', {}).get('credit_points', ''),
            result.get('foundation_course_practical', {}).get('subject_name', ''),
            result.get('foundation_course_practical', {}).get('main_marks', ''),
            result.get('foundation_course_practical', {}).get('cce_marks', ''),
            result.get('foundation_course_practical', {}).get('total_marks', ''),
            result.get('foundation_course_practical', {}).get('max_credits', ''),
            result.get('foundation_course_practical', {}).get('grade', ''),
            result.get('foundation_course_practical', {}).get('grade_points', ''),
            result.get('foundation_course_practical', {}).get('earned_credits', ''),
            result.get('foundation_course_practical', {}).get('credit_points', ''),
            # Major 1
            result.get('major_1', {}).get('subject_name', ''),
            result.get('major_1', {}).get('main_marks', ''),
            result.get('major_1', {}).get('cce_marks', ''),
            result.get('major_1', {}).get('total_marks', ''),
            result.get('major_1', {}).get('max_credits', ''),
            result.get('major_1', {}).get('grade', ''),
            result.get('major_1', {}).get('grade_points', ''),
            result.get('major_1', {}).get('earned_credits', ''),
            result.get('major_1', {}).get('credit_points', ''),
            result.get('major_1_practical', {}).get('subject_name', ''),
            result.get('major_1_practical', {}).get('main_marks', ''),
            result.get('major_1_practical', {}).get('cce_marks', ''),
            result.get('major_1_practical', {}).get('total_marks', ''),
            result.get('major_1_practical', {}).get('max_credits', ''),
            result.get('major_1_practical', {}).get('grade', ''),
            result.get('major_1_practical', {}).get('grade_points', ''),
            result.get('major_1_practical', {}).get('earned_credits', ''),
            result.get('major_1_practical', {}).get('credit_points', ''),
            # Major 2
            result.get('major_2', {}).get('subject_name', ''),
            result.get('major_2', {}).get('main_marks', ''),
            result.get('major_2', {}).get('cce_marks', ''),
            result.get('major_2', {}).get('total_marks', ''),
            result.get('major_2', {}).get('max_credits', ''),
            result.get('major_2', {}).get('grade', ''),
            result.get('major_2', {}).get('grade_points', ''),
            result.get('major_2', {}).get('earned_credits', ''),
            result.get('major_2', {}).get('credit_points', ''),
            result.get('major_2_practical', {}).get('subject_name', ''),
            result.get('major_2_practical', {}).get('main_marks', ''),
            result.get('major_2_practical', {}).get('cce_marks', ''),
            result.get('major_2_practical', {}).get('total_marks', ''),
            result.get('major_2_practical', {}).get('max_credits', ''),
            result.get('major_2_practical', {}).get('grade', ''),
            result.get('major_2_practical', {}).get('grade_points', ''),
            result.get('major_2_practical', {}).get('earned_credits', ''),
            result.get('major_2_practical', {}).get('credit_points', ''),
            # Minor
            result.get('minor', {}).get('subject_name', ''),
            result.get('minor', {}).get('main_marks', ''),
            result.get('minor', {}).get('cce_marks', ''),
            result.get('minor', {}).get('total_marks', ''),
            result.get('minor', {}).get('max_credits', ''),
            result.get('minor', {}).get('grade', ''),
            result.get('minor', {}).get('grade_points', ''),
            result.get('minor', {}).get('earned_credits', ''),
            result.get('minor', {}).get('credit_points', ''),
            result.get('minor_practical', {}).get('subject_name', ''),
            result.get('minor_practical', {}).get('main_marks', ''),
            result.get('minor_practical', {}).get('cce_marks', ''),
            result.get('minor_practical', {}).get('total_marks', ''),
            result.get('minor_practical', {}).get('max_credits', ''),
            result.get('minor_practical', {}).get('grade', ''),
            result.get('minor_practical', {}).get('grade_points', ''),
            result.get('minor_practical', {}).get('earned_credits', ''),
            result.get('minor_practical', {}).get('credit_points', ''),
            # Open
            result.get('open', {}).get('subject_name', ''),
            result.get('open', {}).get('main_marks', ''),
            result.get('open', {}).get('cce_marks', ''),
            result.get('open', {}).get('total_marks', ''),
            result.get('open', {}).get('max_credits', ''),
            result.get('open', {}).get('grade', ''),
            result.get('open', {}).get('grade_points', ''),
            result.get('open', {}).get('earned_credits', ''),
            result.get('open', {}).get('credit_points', ''),
            result.get('open_practical', {}).get('subject_name', ''),
            result.get('open_practical', {}).get('main_marks', ''),
            result.get('open_practical', {}).get('cce_marks', ''),
            result.get('open_practical', {}).get('total_marks', ''),
            result.get('open_practical', {}).get('max_credits', ''),
            result.get('open_practical', {}).get('grade', ''),
            result.get('open_practical', {}).get('grade_points', ''),
            result.get('open_practical', {}).get('earned_credits', ''),
            result.get('open_practical', {}).get('credit_points', ''),
            # VOC
            result.get('voc', {}).get('subject_name', ''),
            result.get('voc', {}).get('main_marks', ''),
            result.get('voc', {}).get('cce_marks', ''),
            result.get('voc', {}).get('total_marks', ''),
            result.get('voc', {}).get('max_credits', ''),
            result.get('voc', {}).get('grade', ''),
            result.get('voc', {}).get('grade_points', ''),
            result.get('voc', {}).get('earned_credits', ''),
            result.get('voc', {}).get('credit_points', ''),
            result.get('voc_practical', {}).get('subject_name', ''),
            result.get('voc_practical', {}).get('main_marks', ''),
            result.get('voc_practical', {}).get('cce_marks', ''),
            result.get('voc_practical', {}).get('total_marks', ''),
            result.get('voc_practical', {}).get('max_credits', ''),
            result.get('voc_practical', {}).get('grade', ''),
            result.get('voc_practical', {}).get('grade_points', ''),
            result.get('voc_practical', {}).get('earned_credits', ''),
            result.get('voc_practical', {}).get('credit_points', ''),
            # Project
            result.get('project/internship', {}).get('subject_name', ''),
            result.get('project/internship', {}).get('main_marks', ''),
            result.get('project/internship', {}).get('cce_marks', ''),
            result.get('project/internship', {}).get('total_marks', ''),
            result.get('project/internship', {}).get('max_credits', ''),
            result.get('project/internship', {}).get('grade', ''),
            result.get('project/internship', {}).get('grade_points', ''),
            result.get('project/internship', {}).get('earned_credits', ''),
            result.get('project/internship', {}).get('credit_points', ''),
            result.get('project/internship_practical', {}).get('subject_name', ''),
            result.get('project/internship_practical', {}).get('main_marks', ''),
            result.get('project/internship_practical', {}).get('cce_marks', ''),
            result.get('project/internship_practical', {}).get('total_marks', ''),
            result.get('project/internship_practical', {}).get('max_credits', ''),
            result.get('project/internship_practical', {}).get('grade', ''),
            result.get('project/internship_practical', {}).get('grade_points', ''),
            result.get('project/internship_practical', {}).get('earned_credits', ''),
            result.get('project/internship_practical', {}).get('credit_points', ''),
            # Totals
            result.get('total_max_credits', ''),
            result.get('total_credits_obtained', ''),
            result.get('total_credit_points', ''),
            result.get('result', ''),
            result.get('category', '')
        ]
        ws.append(row)

    # Save the workbook
    wb.save(file_name)

    logging.info(f"Results written to {file_name}")

# Example usage:
final_results = process_course_results(course_details_id)
write_to_excel(final_results, "/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/FINAL_RESULTS.xlsx")