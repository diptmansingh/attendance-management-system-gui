# Excel-Based Attendance Management System (GUI)

## Description
A desktop graphical user interface (GUI) application built with Python and Tkinter for managing student attendance. Instead of relying on a traditional SQL database, this application cleverly utilizes the `pandas` library to read, update, and save attendance records directly to a local Excel workbook (`.xlsx`). 

The system features two distinct portals: a Student Dashboard for checking attendance status, and a protected Faculty Admin Panel for marking and updating daily records.

## Key Features
* **Student Dashboard:** Students can enter their roll number to view their attendance percentage across all courses.
* **Smart Analytics:** Automatically calculates a "Bunk Budget" (how many classes a student can safely miss) or a "Catch-Up" metric (how many consecutive classes they must attend to reach the minimum required percentage).
* **Faculty Admin Panel:** A dedicated interface for faculty to select courses, choose dates, and mark attendance quickly by double-clicking rows (toggling Present/Absent).
* **Excel Backend:** Data is stored in `attendance_data.xlsx`, making the database highly portable, easily auditable, and simple to back up.

## Prerequisites
* Python 3.x installed on your machine.
* Standard Python libraries: `tkinter` (usually included with standard Python installations).

## Setup and Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/diptmansingh/attendance-management-system-gui.git](https://github.com/diptmansingh/attendance-management-system-gui.git)
   cd attendance-management-system-gui
   ```

2. **Install required dependencies:**
   It is recommended to use a virtual environment. Install the required `pandas` and `openpyxl` libraries using:
   ```bash
   pip install -r requirements.txt
   ```

3. **Database Initialization (Run this first!):**
   The application requires a specifically structured Excel file (`attendance_data.xlsx`) to act as its database. Run the included generator script to instantly create a working template populated with sample courses and students:
   ```bash
   python generate_template.py
   ```

4. **Run the Application:**
   Once the Excel database is generated, you can launch the main GUI:
   ```bash
   python ams.py
   ```
   *(To test the student dashboard immediately, try entering the sample roll number: **ENR001**)*

## Future Enhancements
* **Authentication:** Add password protection to the Faculty Admin Panel to prevent unauthorized access.
* **Data Visualizations:** Integrate `matplotlib` to show attendance trends graphically over the semester.
* **Automated Excel Generation:** Add a startup script that automatically generates a blank template of `attendance_data.xlsx` natively if the file doesn't already exist on the user's machine.

---
*Developed with Python, Tkinter, and Pandas.*
