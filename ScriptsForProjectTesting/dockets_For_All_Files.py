"""
Docket Creation Script
----------------------

Overview:
---------
Creates dockets for all the subjects that are there.
This script generates a seating plan (docket) based on student roll numbers filtered by subject type.
It reads an Excel file containing student details, filters based on the specified subject, and distributes the 
roll numbers into a specified number of columns. The results are saved in a structured folder format.

Requirements:
-------------
The student complete data should be taken out first and must already be present in the excel format. Here no interaction with the database or anything.

Configurable Variables:
-----------------------
1. BASE_PATH: The base directory where the input Excel file is located and where output will be saved.
2. file_name: The name of the Excel file (located inside BASE_PATH) that contains student data.
3. Set the columns config -> orientation, x_columns and subject_type.

Usage:
------
1. Configure the variables: BASE_PATH, file_name, subject_type, subject_name, x_columns.
2. Call the `create_docket` function to generate the seating plan.
"""

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


def create_dockets_for_multiple_columns(file_name, column_configs):
    # Load Excel file
    file_path = os.path.join(BASE_PATH, file_name)
    df = pd.read_excel(file_path)

    # Iterate through the list of column configurations
    for config in column_configs:
        column_name = config.get('column_name')
        x_columns = config.get('x_columns')
        orientation = config.get('orientation')

        # Get unique values in the current column
        unique_values = df[column_name].unique()

        # For each unique value, create a docket
        for subject_name in unique_values:
            create_docket(file_name, column_name, subject_name, x_columns, orientation)


# Example usage
if __name__ == "__main__":
    # Excel file name inside the base path
    file_name = 'student_data.xlsx'
    
    # List of configurations for each column
    column_configs = [
        {
            'column_name': 'foundation_course',  # Column to filter by
            'x_columns': 4,                 # Number of columns in the docket
            'orientation': 'horizontal'     # Orientation of the docket
        },
        {
            'column_name': 'major_1',
            'x_columns': 3,
            'orientation': 'vertical'
        },
        {
            'column_name': 'major_2',
            'x_columns': 5,
            'orientation': 'horizontal'
        }
    ]

    # Create dockets for all specified column configurations
    create_dockets_for_multiple_columns(file_name, column_configs)
