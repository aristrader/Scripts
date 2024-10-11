import pandas as pd
import random

# Define the number of rows you want
num_rows = 100  # You can change this value to 200, 1000, etc.

# Sample data for names (400 first names, 200 last names)
first_names = [
    'Amit', 'Raj', 'Priya', 'Sneha', 'Rahul', 'Pooja', 'Vijay', 'Rani', 'Rohit', 'Neha',
    'Siddharth', 'Anjali', 'Karan', 'Shivam', 'Deepika', 'Nikita', 'Arjun', 'Sakshi', 'Mohit', 'Tanya',
    'Akshay', 'Isha', 'Rakesh', 'Riya', 'Saurabh', 'Simran', 'Abhishek', 'Divya', 'Siddhi', 'Kavya',
    'Vikas', 'Nishant', 'Shreya', 'Vishal', 'Bhavna', 'Sonam', 'Gaurav', 'Snehal', 'Tanvi', 'Dev',
    'Ankit', 'Neeraj', 'Nikhil', 'Rohan', 'Prateek', 'Khushi', 'Manoj', 'Swati', 'Parth', 'Ravi',
    'Shweta', 'Sanjay', 'Komal', 'Sonia', 'Preeti', 'Kajal', 'Nehal', 'Shivangi', 'Nitin', 'Himanshu',
    'Sanjana', 'Aishwarya', 'Kriti', 'Sakshi', 'Madhav', 'Ishaan', 'Devansh', 'Niranjan', 'Utkarsh', 'Anmol',
    'Payal', 'Nimisha', 'Rachit', 'Vivek', 'Tanisha', 'Kartik', 'Aarav', 'Avantika', 'Prachi', 'Lavanya',
    'Vasudha', 'Ananya', 'Kaira', 'Siddhi', 'Devika', 'Yash', 'Akira', 'Nandini', 'Lavanya', 'Shivani',
    'Krishna', 'Navya', 'Rehaan', 'Rishabh', 'Ritika', 'Vaishali', 'Raghav', 'Devraj', 'Tanishq', 'Siddharth',
    'Abhijit', 'Vatsal', 'Parinita', 'Kavisha', 'Pranay', 'Shruti', 'Geetika', 'Bhavika', 'Tanvi', 'Sukhdeep',
    'Aditya', 'Manvi', 'Vaibhav', 'Sonal', 'Sidhant', 'Kanishk', 'Aarvi', 'Anoushka', 'Mitali', 'Chaitanya',
    'Shrishti', 'Bhavesh', 'Kunal', 'Ishita', 'Kashish', 'Simar', 'Saksham', 'Parthiv', 'Jaya', 'Ritika',
    'Tarun', 'Shivendra', 'Krati', 'Neeraj', 'Anirudh', 'Tanmay', 'Yashasvi', 'Bharat', 'Anirudh', 'Advik',
    'Krish', 'Dhruv', 'Kavish', 'Shaan', 'Nandkishor', 'Aayushi', 'Hrithik', 'Samarth', 'Vedant', 'Eshaan',
    'Prabhu', 'Uttara', 'Aarav', 'Darshan', 'Reva', 'Shashank', 'Vinay', 'Anay', 'Parmeet', 'Yug',
    'Eshaan', 'Tejas', 'Vaidik', 'Shailesh', 'Pankaj', 'Prathamesh', 'Rupal', 'Ruchika', 'Rhea', 'Samiksha',
    'Zeenat', 'Sharmila', 'Falguni', 'Radhika', 'Ambika', 'Vaishnavi', 'Kumud', 'Rukmini', 'Rani', 'Bela',
    'Anupama', 'Urmila', 'Rashmi', 'Aditi', 'Jyoti', 'Vani', 'Kamal', 'Vijayalaxmi', 'Chhavi', 'Sushma',
    'Archana', 'Poonam', 'Deepa', 'Sonali', 'Gurpreet', 'Gargi', 'Chhaya', 'Chandni', 'Aakriti', 'Sangeeta',
    'Leela', 'Poonam', 'Meenakshi', 'Usha', 'Saloni', 'Madhuri', 'Savita', 'Kakoli', 'Rashmi', 'Kaveri',
    'Mridula', 'Shilpa', 'Kamini', 'Naina', 'Suman', 'Geeta', 'Malti', 'Vijaya', 'Devaki', 'Aruna',
    'Radha', 'Sudha', 'Ankita', 'Deepali', 'Namrata', 'Sampada', 'Diksha', 'Neha', 'Kanta', 'Jayashree',
    'Veena', 'Vidya', 'Chitra', 'Shanti', 'Suman', 'Bharti', 'Dharini', 'Medha', 'Gina', 'Priyanka',
    'Sahiba', 'Ritika', 'Bhanita', 'Nidhi', 'Sharanya', 'Ankita', 'Kiran', 'Rekha', 'Sushmita', 'Neeti',
    'Sakshi', 'Jayanti', 'Kartikeya', 'Sidhartha', 'Gaurika', 'Sukhbir', 'Deependra', 'Himadri', 'Jaya', 
    'Kaushal', 'Rekha', 'Asha', 'Bhanu', 'Rukmani', 'Tanisha', 'Shailaja', 'Salina', 'Ankush', 'Shivendra',
    'Tanuja', 'Riya', 'Pranita', 'Rajendra', 'Shanaya', 'Rajesh', 'Dharma', 'Madhura', 'Ashutosh', 'Arpita',
    'Kashika', 'Neerav', 'Kanishka', 'Sandeep', 'Vaibhav', 'Nirav', 'Vidur', 'Ira', 'Nishtha', 'Vaibhav',
    'Viraj', 'Mahesh', 'Naman', 'Siddhart', 'Yogesh', 'Lokesh', 'Siddhant', 'Aarush', 'Harsh', 'Ashish',
    'Anshika', 'Aniketh', 'Ajinkya', 'Rajvir', 'Kashvi', 'Avyan', 'Charvi', 'Yashvi', 'Manan', 'Parag'
]

last_names = [
    'Sharma', 'Gupta', 'Mehta', 'Kumar', 'Singh', 'Verma', 'Kapoor', 'Mishra', 'Patel', 'Yadav',
    'Chauhan', 'Sinha', 'Bhatia', 'Jain', 'Malhotra', 'Aggarwal', 'Shukla', 'Saxena', 'Ghosh', 'Bansal',
    'Ahuja', 'Choudhary', 'Sethi', 'Khan', 'Desai', 'Rai', 'Nair', 'Jha', 'Pandey', 'Yadav',
    'Choudhary', 'Sood', 'Dutt', 'Khatri', 'Reddy', 'Nambiar', 'Menon', 'George', 'Philip', 'John',
    'Kumar', 'Mishra', 'Pandit', 'Kapoor', 'Tyagi', 'Nair', 'Tiwari', 'Kumar', 'Khandekar', 'Suryavanshi',
    'Srinivasan', 'Joshi', 'Dey', 'Dutta', 'Bhattacharya', 'Sen', 'Ghosh', 'Basu', 'Bhattacharya', 'Deb',
    'Roy', 'Mukherjee', 'Chowdhury', 'Chatterjee', 'Banerjee', 'Chakraborty', 'Saha', 'Dutta', 'Poddar', 'Khan',
    'Rathore', 'Bhatnagar', 'Chopra', 'Hegde', 'Kaul', 'Nath', 'Gupta', 'Jain', 'Puri', 'Vora',
    'Gandhi', 'Mishra', 'Kothari', 'Rathore', 'Kapoor', 'Kohli', 'Kachroo', 'Parmar', 'Chandel', 'Bishnoi',
    'Kumar', 'Raina', 'Sharma', 'Bhatnagar', 'Malik', 'Rana', 'Gurjar', 'Kakkar', 'Sethi', 'Bansal',
    'Nag', 'Talwar', 'Kumar', 'Mehta', 'Nanda', 'Dandekar', 'Kothari', 'Murlidhar', 'Bose', 'Bhandari',
    'Kumar', 'Mishra', 'Dhillon', 'Chahal', 'Suri', 'Saigal', 'Chawla', 'Bhagwat', 'Prakash', 'Singhal',
    'Anand', 'Dhamija', 'Handa', 'Agarwal', 'Chaturvedi', 'Jaiswal', 'Kaur', 'Nanda', 'Kohli', 'Tanwar',
    'Vashisht', 'Awasthi', 'Awasthi', 'Khanna', 'Khan', 'Talwar', 'Hirawat', 'Bhat', 'Jindal', 'Sarkar',
    'Surana', 'Sharma', 'Choudhary', 'Nirula', 'Pandit', 'Pandey', 'Datt', 'Satyavrat', 'Bansal', 'Bhatti',
    'Thakur', 'Saini', 'Shankar', 'Ghoshal', 'Dabas', 'Dar', 'Kohli', 'Mokha', 'Singh', 'Arora',
    'Kaushik', 'Jangra', 'Ajwani', 'Sastri', 'Vishwakarma', 'Kesarwani', 'Mahajan', 'Sethi', 'Bharadwaj', 'Chopra',
    'Vardhan', 'Khunteta', 'Tanwar', 'Satyam', 'Jangra', 'Gupta', 'Nehra', 'Chaudhary', 'Sharma', 'Bhargav',
    'Lal', 'Bansal', 'Shrivastava', 'Rao', 'Varma', 'Bhagwati', 'Kakkar', 'Malhotra', 'Kapoor', 'Khanna',
    'Jindal', 'Sengupta', 'Malik', 'Rajput', 'Thakur', 'Kharbanda', 'Chand', 'Dadhich', 'Choudhary', 'Rawat'
]

# Define unique subject choices
subject_choices = {
    'foundation_course': [
        'History', 'Geography', 'Physics', 'Sociology'
    ],
    
    'major_1': [
        'Psychology', 'Computer Science', 'Chemistry'
    ],
    
    'major_2': [
        'Economics', 'Political Science', 'Biology'
    ],
    
    'minor': [
        'English', 'Accounting', 'Marketing', 'Banking'
    ],
    
    'open': [
        'Maths', 'Tourism', 'Event Management', 'Fashion Design'
    ],
    
    'voc': [
        'Retail Management', 'Graphic Design', 'IT', 'Photography'
    ],
    
    'project_internship': ['Project', 'Internship']  # Fixed options
}

# Helper function to generate random names
def generate_name():
    return random.choice(first_names) + ' ' + random.choice(last_names)

# Helper function to generate random roll number and enrollment number
def generate_roll_no():
    return random.randint(10000, 99999)

def generate_enrollment_no():
    return random.randint(1000000, 9999999)

# Generate the data
data = []
for i in range(1, num_rows + 1):
    row = {
        's.no': i,
        'roll no': generate_roll_no(),
        'enrollment no': generate_enrollment_no(),
        'student name': generate_name(),
        'father name': generate_name(),
        'mother name': generate_name(),
        'foundation_course': random.choice(subject_choices['foundation_course']),
        'major_1': random.choice(subject_choices['major_1']),
        'major_2': random.choice(subject_choices['major_2']),
        'minor': random.choice(subject_choices['minor']),
        'open': random.choice(subject_choices['open']),
        'voc': random.choice(subject_choices['voc']),
        'project/internship': random.choice(subject_choices['project_internship'])
    }
    data.append(row)

# Convert to pandas DataFrame
df = pd.DataFrame(data)

# Save to Excel
output_file = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels/student_data.xlsx'
df.to_excel(output_file, index=False)

print(f"Excel file with {num_rows} rows has been created: {output_file}")
