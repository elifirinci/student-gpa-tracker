# Halic University Student GPA Tracker

This is a Python-based desktop application developed using the `tkinter` library for graphical user interfaces. The application allows users to load an Excel file containing student records, calculate GPAs and rankings, and display or export specific student information based on their ID.

## Features
- **Load Excel Files**: Open and process `.xlsx` files containing student information and grades.
- **Calculate GPA and Rank**: Automatically calculates the GPA and rank for each student based on their grades.
- **Display Student Information**: Retrieve and display student name, surname, GPA, and rank by entering their ID.
- **Export Student Data**: Save the displayed student information to a file (`.txt` or `.xls`) for easy sharing or storage.
- **Clear Data**: Reset the input fields and displayed information for new searches.

## Installation
1. Ensure you have Python 3.x installed on your system.
2. Install the required Python package using the following command:
   ```bash
   pip install openpyxl
## Clone this repository:
git clone https://github.com/elifirinci/student-gpa-tracker.git
## Usage
**Run the application**:
- python student_gpa_tracker.py
- Load an Excel file with student records by clicking the Browse button.
- The file should have the following columns: Name, Surname, Student ID, Physics, Calculus, Programming, Chemistry.
- Enter a student's ID and click Display to see their information.
- Export the displayed information to a file using the Export as File button.
- Use the Clear button to reset the input fields and displayed data.
