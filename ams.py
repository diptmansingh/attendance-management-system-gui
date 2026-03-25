import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import datetime


EXCEL_FILE = "attendance_data.xlsx" 

class FacultyWindow(tk.Toplevel):
    """Window for Faculty to view, mark, and save attendance."""
    
    def __init__(self, master, app_instance):
        super().__init__(master)
        self.title("Faculty Admin Panel")
        self.geometry("1000x650")
        self.app = app_instance
        self.grab_set() 

        self.selected_course = tk.StringVar(self)
        self.selected_date = tk.StringVar(self)
        self.current_df = pd.DataFrame()
        self.current_tab_name = ""

        self.setup_ui()
        self.load_course_options()

    def setup_ui(self):
        
        control_frame = ttk.Frame(self, padding="10")
        control_frame.pack(fill='x', pady=5)
        
        ttk.Label(control_frame, text="Select Course:").pack(side=tk.LEFT, padx=5)
        self.course_menu = ttk.OptionMenu(control_frame, self.selected_course, 'Select Course', command=self.load_attendance_data)
        self.course_menu.pack(side=tk.LEFT, padx=10)
        
        ttk.Label(control_frame, text="Attendance Date (YYYY-MM-DD):").pack(side=tk.LEFT, padx=15)
        
        today = datetime.date.today().strftime('%Y-%m-%d')
        self.selected_date.set(today)
        self.date_entry = ttk.Entry(control_frame, textvariable=self.selected_date, width=12)
        self.date_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Mark Attendance", command=self.load_attendance_data).pack(side=tk.LEFT, padx=15)
        ttk.Button(control_frame, text="Save Changes to Excel", command=self.save_attendance_to_excel, style='Accent.TButton').pack(side=tk.RIGHT, padx=5)

      
        self.status_label = ttk.Label(self, text="Select a course and date to begin marking.", foreground="blue")
        self.status_label.pack(fill='x', padx=10, pady=(0, 5))

        
        self.tree = ttk.Treeview(self, columns=('Roll', 'Name', 'Status'), show='headings')
        self.tree.heading('Roll', text='Roll Number')
        self.tree.heading('Name', text='Student Name')
        self.tree.heading('Status', text='Attendance (Click to Toggle P/A)')
        
        self.tree.column('Roll', width=100, anchor='center')
        self.tree.column('Name', width=200, anchor='w')
        self.tree.column('Status', width=250, anchor='center')
        
        self.tree.pack(fill='both', expand=True, padx=10, pady=10)
        self.tree.bind('<Double-1>', self.toggle_attendance)

    def load_course_options(self):
        """Populates the course selection dropdown."""
        course_names = self.app.course_details_df['Course_Code'].tolist()
        
        
        menu = self.course_menu['menu']
        menu.delete(0, 'end')
        for name in course_names:
            menu.add_command(label=name, command=lambda value=name: self.selected_course.set(value))
        
        if course_names:
            self.selected_course.set(course_names[0])


    def load_attendance_data(self, *args):
        """Loads data for the selected course and populates the Treeview."""
        course_name = self.selected_course.get()
        attendance_date = self.selected_date.get()

        if course_name == 'Select Course' or not attendance_date:
            self.status_label.config(text="Please select a valid course and date.", foreground="red")
            return

        
        try:
            tab_name = self.app.course_details_df[self.app.course_details_df['Course_Code'] == course_name]['Sheet_Tab_Name'].iloc[0]
            self.current_tab_name = tab_name
        except IndexError:
            self.status_label.config(text=f"Error: Course '{course_name}' configuration not found.", foreground="red")
            return

        
        if course_name not in self.app.attendance_dfs:
            self.status_label.config(text=f"Error: Attendance data for '{course_name}' not loaded.", foreground="red")
            return
            
        self.current_df = self.app.attendance_dfs[course_name].copy() 

        
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        
        if attendance_date not in self.current_df.columns:
            self.current_df[attendance_date] = '' 

       
        for index, row in self.current_df.iterrows():
            status = row.get(attendance_date, '')
            
            self.tree.insert('', tk.END, iid=str(index), values=(
                row['Enrollment_Number'],
                row['Student_Name'],
                status if status else 'Not Taken'
            ), tags=(status.lower(),)) 

  
        self.tree.tag_configure('p', background='lightgreen')
        self.tree.tag_configure('a', background='#ffcccb')
        
        self.status_label.config(text=f"Attendance loaded for {course_name} on {attendance_date}. Double-click to mark.", foreground="blue")


    def toggle_attendance(self, event):
        """Toggles the attendance status for the double-clicked row."""
        item_id = self.tree.focus()
        if not item_id:
            return

       
        row_index = int(item_id) 
        current_values = self.tree.item(item_id, 'values')
        
       
        current_status = current_values[2].strip().upper()

        if current_status == 'P':
            new_status = 'A'
        elif current_status == 'A':
            new_status = 'Not Taken'
        else:
            new_status = 'P'

        
        new_values = list(current_values)
        new_values[2] = new_status
        self.tree.item(item_id, values=new_values)
        
       
        attendance_date = self.selected_date.get()
        if new_status == 'Not Taken':
            
            self.current_df.loc[row_index, attendance_date] = ''
        else:
            self.current_df.loc[row_index, attendance_date] = new_status[0]
        
        
        self.tree.item(item_id, tags=(new_status[0].lower() if new_status != 'Not Taken' else '',))

        self.status_label.config(text=f"Status for {current_values[1]} updated to {new_status}. **Don't forget to SAVE!**", foreground="darkgreen")


    def save_attendance_to_excel(self):
        """Saves the modified DataFrame back to the specific sheet in the Excel file."""
        if self.current_df.empty or not self.current_tab_name:
            messagebox.showerror("Save Error", "No attendance data loaded or course selected.")
            return

        try:
            
            with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                
                self.current_df.to_excel(writer, sheet_name=self.current_tab_name, index=False)
            
            messagebox.showinfo("Save Success", f"Attendance for {self.current_tab_name} saved successfully!")
            
            
            self.app.load_excel_data() 

        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save data: {e}. Check if the file is open.")



class AttendanceApp(tk.Tk):

    
    def __init__(self):
        super().__init__()
        self.title("Student Attendance Portal (Excel Based)")
        self.geometry("800x600")
        
        self.course_details_df = pd.DataFrame()
        self.attendance_dfs = {} 
        self.student_enrollment_number = tk.StringVar()

        self.setup_ui()
        self.load_excel_data()

    def load_excel_data(self):
        self.attendance_dfs = {} 
        try:
            excel_data = pd.read_excel(EXCEL_FILE, sheet_name=None)
            
            if 'Course_Details' not in excel_data:
                raise ValueError("Missing 'Course_Details' tab in the Excel file.")
            
            self.course_details_df = excel_data['Course_Details']

            for index, row in self.course_details_df.iterrows():
                tab_name = row['Sheet_Tab_Name']
                course_name = row['Course_Code']
                
                if tab_name not in excel_data:
                    self.update_status(f"Warning: Tab '{tab_name}' listed in Course_Details is missing.", "orange")
                    continue
                
                self.attendance_dfs[course_name] = excel_data[tab_name]
            
            self.update_status(f"Data loaded successfully from {EXCEL_FILE}!", "green")
            
        except FileNotFoundError:
            self.update_status(f"Error: '{EXCEL_FILE}' not found. Check the file path!", "red")
        except Exception as e:
            self.update_status(f"Error loading data: {e}. Check Excel structure.", "red")


    def setup_ui(self):
        """Sets up the Tkinter graphical user interface, including the Faculty button."""
        
        
        input_frame = ttk.Frame(self, padding="10")
        input_frame.pack(fill='x', pady=10)
        
        ttk.Label(input_frame, text="Your Roll Number:", font=('Arial', 12, 'bold')).pack(side=tk.LEFT, padx=5)
        ttk.Entry(input_frame, textvariable=self.student_enrollment_number, width=15, font=('Arial', 12)).pack(side=tk.LEFT, padx=5)
        ttk.Button(input_frame, text="Show Attendance", command=self.display_attendance_report, style='TButton').pack(side=tk.LEFT, padx=15)
        ttk.Button(input_frame, text="Refresh Data (Read Excel)", command=self.load_excel_data, style='TButton').pack(side=tk.LEFT, padx=5)
        
        
        ttk.Button(input_frame, text="Faculty Admin 🔒", command=self.open_faculty_admin, style='Danger.TButton').pack(side=tk.RIGHT, padx=5)


        
        self.status_label = ttk.Label(self, text="Application Ready.", foreground="black")
        self.status_label.pack(fill='x', padx=10, pady=(0, 5))
        
        
        self.report_frame = ttk.Frame(self, padding="10")
        self.report_frame.pack(fill='both', expand=True)

        self.report_label = ttk.Label(self.report_frame, text="Enter your Roll Number to view your report.", font=('Arial', 14))
        self.report_label.pack(pady=20)
        
       
        self.tree = ttk.Treeview(self.report_frame, columns=('Course', 'Attended', 'Held', 'Percent', 'Status'), show='headings')
        self.tree.heading('Course', text='Course Name')
        self.tree.heading('Attended', text='Classes Attended')
        self.tree.heading('Held', text='Classes Held')
        self.tree.heading('Percent', text='Percentage (%)')
        self.tree.heading('Status', text='Status / Action')

        self.tree.column('Course', width=150, anchor='w')
        self.tree.column('Attended', width=100, anchor='center')
        self.tree.column('Held', width=100, anchor='center')
        self.tree.column('Percent', width=100, anchor='center')
        self.tree.column('Status', width=250, anchor='w')

        self.tree.pack(fill='both', expand=True)

    def open_faculty_admin(self):
        """Opens the separate Faculty Administration window."""
        if self.course_details_df.empty:
             messagebox.showerror("Error", "Cannot open Admin Panel. Data not loaded. Check Excel file.")
             return
            
        FacultyWindow(self, self)

    
    def update_status(self, message, color="black"):
        """Updates the status bar message."""
        self.status_label.config(text=message, foreground=color)


    def display_attendance_report(self):
        
        enrollment_number = self.student_enrollment_number.get().strip().upper()
        if not enrollment_number:
            messagebox.showwarning("Input Error", "Please enter your Roll Number.")
            return

        if not self.attendance_dfs:
            messagebox.showerror("Data Error", f"No attendance data loaded. Ensure {EXCEL_FILE} exists.")
            return
            
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        student_found = False
        
        for course_name, attendance_df in self.attendance_dfs.items():

            attendance_df['Enrollment_Number_str'] = attendance_df['Enrollment_Number'].astype(str).str.strip().str.upper()
            
            student_data = attendance_df[attendance_df['Enrollment_Number_str'] == enrollment_number]

            if student_data.empty:
                continue
                
            student_found = True
            
            course_config = self.course_details_df[self.course_details_df['Course_Code'] == course_name].iloc[0]
            min_percent = course_config['Minimum_Percentage']
            student_name = student_data['Student_Name'].iloc[0]
            self.report_label.config(text=f"Report for: {student_name} (Roll: {enrollment_number})")

            row = student_data.iloc[0]
            attendance_columns = [col for col in attendance_df.columns if col not in ['Enrollment_Number', 'Student_Name']]
            attendance_marks = row[attendance_columns].astype(str)
            
            classes_held = attendance_marks[attendance_marks.str.strip() != ''].count()-1
            classes_attended = (attendance_marks.str.strip().str.upper() == 'P').sum()
            
            percentage = (classes_attended / classes_held * 100) if classes_held > 0 else 0
            percentage = round(percentage, 2)
            
            status_text = ""
            if percentage >= min_percent:
                status_text = f"✅ Above required {min_percent}%."
                bunk_budget = self.calculate_bunk_budget(classes_attended, classes_held, min_percent)
                if bunk_budget > 0:
                    status_text += f" (Can miss **{bunk_budget}** more classes)"
            else:
                status_text = f"❌ BELOW {min_percent}%!"
                catch_up_classes = self.calculate_catch_up(classes_attended, classes_held, min_percent)
                status_text += f" (MUST attend next **{catch_up_classes}** classes to reach {min_percent}%)"
            
            self.tree.insert('', tk.END, values=(
                course_name,
                classes_attended,
                classes_held,
                percentage,
                status_text
            ))

        if not student_found:
             messagebox.showerror("Roll Number Error", f"Roll Number '{enrollment_number}' not found in any courses.")


    def calculate_catch_up(self, present, held, target_percent):
        target = target_percent / 100.0
        x = 0
        while True:
            current_ratio = (present + x) / (held + x) if (held + x) > 0 else 0
            if current_ratio >= target: return x
            if x > 100: return "100+ (Review required)"
            x += 1

    def calculate_bunk_budget(self, present, held, target_percent):
        target = target_percent / 100.0
        x = 0
        while True:
            if present - x < 0: return x - 1 if x > 0 else 0 
            current_ratio = (present - x) / (held + x)
            if current_ratio < target: return x - 1 
            if x > 100: return 0
            x += 1

if __name__ == "__main__":
    app = AttendanceApp()
    
    
    style = ttk.Style(app)
    style.theme_use('clam')
    style.configure('TButton', font=('Arial', 10), padding=6)
    style.configure('Accent.TButton', foreground='white', background='#007bff')
    style.map('Accent.TButton', background=[('active', '#0056b3')])
    style.configure('Danger.TButton', foreground='white', background='#dc3545')
    style.map('Danger.TButton', background=[('active', '#c82333')])
    
    app.mainloop()