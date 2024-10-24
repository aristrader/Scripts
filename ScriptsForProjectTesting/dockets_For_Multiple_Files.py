import os
import pandas as pd
from math import ceil

# Step 1 : Global base path for the project
BASE_PATH = '/Users/swapnilagarwal/Visual_Studio_Projects/Results/TestingExcels'

def create_docket(file_name, subject_type, subject_name, x_columns, orientation='horizontal'):
    # Load Excel file
    file_path = os.path.join(BASE_PATH, file_name)
    df = pd.read_excel(file_path)

    # Filter rows based on subject type and subject name
    filtered_df = df[df[subject_type] == subject_name].copy()  # Use .copy() to avoid SettingWithCopyWarning

    # Convert roll_no to integers for sorting (handles strings with numeric content)
    filtered_df.loc[:, 'roll no'] = filtered_df['roll no'].astype(int)

    # Sort the roll numbers numerically
    sorted_roll_numbers = filtered_df['roll no'].sort_values()

    # Distribute roll numbers based on orientation
    if orientation == 'horizontal':
        roll_numbers_split = [sorted_roll_numbers[i::x_columns].reset_index(drop=True) for i in range(x_columns)]
        
        # Find the maximum length for padding
        max_length = max(len(col) for col in roll_numbers_split)
        
        # Pad shorter columns with NaN
        for i in range(len(roll_numbers_split)):
            roll_numbers_split[i] = roll_numbers_split[i].reindex(range(max_length))

        # Create a new DataFrame for the seating plan
        seating_plan_df = pd.DataFrame({f'Column_{i+1}': col.values for i, col in enumerate(roll_numbers_split)})

    elif orientation == 'vertical':
        num_roll_numbers = len(sorted_roll_numbers)
        # Calculate how many rows will be needed
        num_rows = ceil(num_roll_numbers / x_columns)
        
        # Create a list to hold the roll numbers in the vertical order
        vertical_list = [[] for _ in range(num_rows)]
        
        for i in range(num_roll_numbers):
            vertical_list[i % num_rows].append(sorted_roll_numbers.iloc[i])
        
        # Create a DataFrame with the specified number of columns
        seating_plan_df = pd.DataFrame({f'Column_{i+1}': [row[i] if i < len(row) else None for row in vertical_list] for i in range(num_rows)})

        # Fill in the remaining columns with NaN if necessary
        for i in range(num_rows):
            while len(vertical_list[i]) < x_columns:
                vertical_list[i].append(None)

        # Create final seating plan DataFrame
        seating_plan_df = pd.DataFrame({f'Column_{i+1}': [vertical_list[j][i] if j < len(vertical_list) and i < len(vertical_list[j]) else None for j in range(num_rows)] for i in range(x_columns)})

    else:
        raise ValueError("Invalid orientation. Choose 'horizontal' or 'vertical'.")

    # Create the folder structure for saving the output
    output_folder = os.path.join(BASE_PATH, 'Dockets', subject_type)
    os.makedirs(output_folder, exist_ok=True)

    # Save the seating plan as an Excel file
    output_file_path = os.path.join(output_folder, f'{subject_name}.xlsx')
    seating_plan_df.to_excel(output_file_path, index=False)

    print(f"Docket created and saved at: {output_file_path}")


def create_dockets_for_multiple_subjects(file_name, subject_configs):
    for config in subject_configs:
        subject_type = config.get('subject_type')
        subject_name = config.get('subject_name')
        columns = config.get('columns')
        orientation = config.get('orientation')

        create_docket(file_name, subject_type, subject_name, columns, orientation)


# Example usage
if __name__ == "__main__":
    # Excel file name inside the base path
    file_name = 'student_data.xlsx'
    
    # List of configurations for subject types, names, columns, and orientations
    subject_configs = [
        {
            'subject_type': 'foundation_course',
            'subject_name': 'Physics',
            'columns': 5,
            'orientation': 'vertical'
        },
        {
            'subject_type': 'foundation_course',
            'subject_name': 'Geography',
            'columns': 4,
            'orientation': 'horizontal'
        },
        {
            'subject_type': 'major_2',
            'subject_name': 'Biology',
            'columns': 3,
            'orientation': 'vertical'
        },
        {
            'subject_type': 'major_2',
            'subject_name': 'Economics',
            'columns': 2,
            'orientation': 'horizontal'
        }
    ]

    # Create dockets for the specified subjects
    create_dockets_for_multiple_subjects(file_name, subject_configs)
