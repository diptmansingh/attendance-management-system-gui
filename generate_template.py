import pandas as pd

def create_template():
    filename = "attendance_data.xlsx"

    # 1. Create the Course_Details configuration sheet
    course_details_data = {
        "Course_Code": ["DBMS201", "MTH202", "CYB301"],
        "Sheet_Tab_Name": ["DBMS_Attendance", "Math_Attendance", "Cyber_Attendance"],
        "Minimum_Percentage": [75, 75, 80]
    }
    df_courses = pd.DataFrame(course_details_data)

    # 2. Create sample student rosters for each course
    dbms_students = pd.DataFrame({
        "Enrollment_Number": ["ENR001", "ENR002", "ENR003"],
        "Student_Name": ["Alice Smith", "Bob Jones", "Charlie Brown"]
    })

    math_students = pd.DataFrame({
        "Enrollment_Number": ["ENR001", "ENR002", "ENR004"],
        "Student_Name": ["Alice Smith", "Bob Jones", "Diana Prince"]
    })
    
    cyber_students = pd.DataFrame({
        "Enrollment_Number": ["ENR001", "ENR003", "ENR004"],
        "Student_Name": ["Alice Smith", "Charlie Brown", "Diana Prince"]
    })

    # 3. Write everything to an Excel workbook
    print(f"Generating {filename}...")
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df_courses.to_excel(writer, sheet_name="Course_Details", index=False)
            dbms_students.to_excel(writer, sheet_name="DBMS_Attendance", index=False)
            math_students.to_excel(writer, sheet_name="Math_Attendance", index=False)
            cyber_students.to_excel(writer, sheet_name="Cyber_Attendance", index=False)
            
        print("✅ Template generated successfully! You can now run your main application.")
    except Exception as e:
        print(f"❌ Error generating template: {e}")

if __name__ == "__main__":
    create_template()