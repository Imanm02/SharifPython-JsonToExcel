# Import necessary libraries
import pandas as pd
import json
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# Load data from json file
# This is where you read the data from your JSON file into a Python dictionary.
with open("9090.json", 'r', encoding='utf-8') as f:
    data = json.load(f)

# Define a dictionary to map the values of 'internal_messenger_type'
# You can add or remove mappings according to your needs.
messenger_type_mapping = {
    'bale': 'بله',
    'eitaa': 'ایتا',
    'rubika': 'روبیکا',
    'gap': 'گپ',
    'soroush': 'سروش',
    'igap': 'آی‌گپ'
}

# Initialize an empty list to store student data
students_data = []

# Loop over each course in the data
for course in data['courses']:
    # Loop over each group in the current course
    for group in course['current_group']:
        # Loop over each student in the current group
        for student in group['students']:
            # Extract the relevant information from the student dictionary
            # and store it in a new dictionary with the desired keys (column names)
            # Apply the necessary transformations to the values
            student_data = {
                'نام': student['name'],
                'نام خانوادگی': student['surname'],
                'جنسیت': 'آقا' if student['gender'] == 'male' else 'خانم',
                'ایمیل': student['email'],
                'شماره تلفن': student['mobile_number'],
                'کد ملی': student['national_code'],
                'شماره تلفن ثابت': student['phone_number'],
                'تاریخ آپدیت اطلاعات': student['updated_at'],
                'وضعیت': 'ثبت‌نام' if student['pivot']['status'] == 'pending' else 'کردیت',
            }

            # Check if the 'extra' field contains a dictionary
            if student['extra'] and isinstance(student['extra'], dict):
                # Loop over each key-value pair in the 'extra' dictionary
                for key, value in student['extra'].items():
                    # Handle each key according to its name
                    if key == 'telegram_number':
                        student_data['شماره تلگرام'] = value
                    elif key == 'whatsapp_number':
                        student_data['شماره واتس‌آپ'] = value
                    elif key == 'internal_messenger_type':
                        # Use the mapping defined earlier to transform the value
                        student_data['پیام‌رسان داخلی مورد استفاده'] = messenger_type_mapping.get(value, value)
                    elif key == 'internal_messenger_number':
                        student_data['شماره تلفن پیام‌رسان داخلی'] = value
                    elif key == 'in_person_classes':
                        student_data['وضعیت حضور در کلاس'] = 'حضوری' if value == True else 'مجازی'

            # Add the student's data to the list
            students_data.append(student_data)

# Create a dataframe from the data
df = pd.DataFrame(students_data)

# Write the dataframe to an Excel file
df.to_excel('9090.xlsx', index=False)

# Load the Excel file back into memory
wb = openpyxl.load_workbook('9090.xlsx')
sheet = wb.active

# Define the styles to apply to the cells
font = Font(name='Vazirmatn')
alignment = Alignment(horizontal='center', vertical='center')
light_yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
light_green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

# Apply the styles to the cells
for row in sheet:
    for cell in row:
        cell.font = font
        cell.alignment = alignment
        cell.number_format = '@'  # Set number format to text
        # Apply different background colors to the first row and the rest
        if cell.row == 1:
            cell.fill = light_yellow_fill
        else:
            cell.fill = light_green_fill

# Adjust the width of the columns
for column_cells in sheet.columns:
    length = max(len(str(cell.value)) for cell in column_cells)
    sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length

# Save the changes made to the Excel file
wb.save('9090.xlsx')