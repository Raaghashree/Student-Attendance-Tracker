import openpyxl  

file_path = 'D:\\attendance.xlsx'
book = openpyxl.load_workbook(uploaded)
sheet = book['Sheet1']

def save_file():
    book.save(file_path)
    print("Attendance record updated successfully!")

def update_attendance(subject_col, roll_numbers):
    for roll_no in roll_numbers:
        for row in range(2, sheet.max_row + 1):
            if sheet.cell(row=row, column=1).value == roll_no:
                sheet.cell(row=row, column=subject_col).value += 1
                break
    save_file()

while True:
    print("Subjects:\n1 -> CI\n2 -> Python\n3 -> DM")
    subject_choice = int(input("Enter subject number: "))

    if subject_choice not in [1, 2, 3]:
        print("Invalid choice. Try again.")
        continue

    subject_col = 2 + subject_choice 
    num_absentees = int(input("Enter number of absentees: "))

    roll_numbers = list(map(int, input("Enter roll numbers (space-separated): ").split())) if num_absentees > 1 else [int(input("Enter roll number: "))]

    update_attendance(subject_col, roll_numbers)

    if int(input("Update another subject? 1 -> Yes, 0 -> No: ")) == 0:
        break
