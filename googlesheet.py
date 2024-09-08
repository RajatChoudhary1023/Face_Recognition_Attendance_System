import gspread
import random

# Authenticate and open the spreadsheet
gc = gspread.service_account(filename="C:\\Users\\sansk\\Downloads\\ast-auto-5041a107d45e.json")
worksheet = gc.open_by_key("1zxHrXYcewYXA7MWkhVykOHL_3rQIBs9I7yL-CkwS4Sc")
current_sheet = worksheet.sheet1

# Fetch all values from the sheet
cell_values = current_sheet.get_all_values()

# Iterate over each row, starting from the second row
for row_idx in range(1, len(cell_values)):
    row = cell_values[row_idx]
    roll_no = row[0]

    try:
        # Extract the numeric part of the roll number
        roll_number = int(roll_no[2:])  # Assumes roll numbers are in format "TC" followed by digits
        
        # Decide attendance based on whether the roll number is even or odd
        if roll_number % 2 == 0:
            status = "Present"
        else:
            status = "Absent"
        
        # Update the 'attendance' column (which is the 3rd column, index 2)
        current_sheet.update_cell(row_idx + 1, 3, status)
    
    except ValueError:
        # If the roll number cannot be converted to an integer, skip this row
        print(f"Skipping invalid roll number: {roll_no}")

print("Attendance updated successfully.")
