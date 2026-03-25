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
* Standard Python libraries: `tkinter` (usually included with Python).

## Setup and Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/diptmansingh/your-repo-name.git](https://github.com/diptmansingh/your-repo-name.git)
   cd your-repo-name
   ```

2. **Install required dependencies:**
   It is recommended to use a virtual environment. Install the required pandas and openpyxl libraries using:
   ```bash
   pip install -r requirements.txt
   ```

3. **Database Configuration (Crucial Step):**
   The application requires an Excel file named `attendance_data.xlsx` in the root directory. It must contain the following structure:
   * A sheet named exactly **`Course_Details`** with columns: `Course_Code`, `Sheet_Tab_Name`, and `Minimum_Percentage`.
   * Individual sheets corresponding to the `Sheet_Tab_Name` values. These sheets must contain at least two columns: `Enrollment_Number` and `Student_Name`. 

4. **Run the Application:**
   ```bash
   python ams.py
   ```

## Future Enhancements
* **Authentication:** Add actual password protection to the Faculty Admin Panel.
* **Data Visualizations:** Integrate `matplotlib` to show attendance trends graphically.
* **Automated Excel Generation:** Add a startup script that automatically generates a blank template of `attendance_data.xlsx` if the file doesn't already exist.

---
*Developed with Python, Tkinter, and Pandas.*
