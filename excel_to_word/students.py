import numpy as np
import pandas as pd


required_headers = [
    "TA Name",
    "Student Number",
    "Email",
    "Subject",
    "Course Code",
    "Sec No.",
    "Act Type",
    "Days Met",
    "Start Time",
    "End Time",
    "TA Hours",
    "Instructor",
    "Position",
    "Wage/month",
    "Total Hours"
]


class Student:

    def __init__(self, name, student_id, position,
                 salary, hours_per_semester, email=None):

        self.name = name
        self.email = email
        self.student_id = student_id
        self.assignments = []
        self.duty_hours = 0
        self.total_hours_per_semester = hours_per_semester
        self.position = position
        self.salary = salary
        self.assigned_courses = {}


class Assignment:
    def __init__(self, subject, course, section, type,
                 days_met, start_time, end_time, hours):

        self.subject = subject
        self.course = course
        self.section = section
        self.type = type
        self.days_met = days_met
        self.start_time = start_time
        self.end_time = end_time
        self.hours = hours


class TA:
    def __init__(self, path):
        self.data_frame = None
        self.students = list()
        self.read_data(path)
        self.create_student()

    def read_data(self, path):
        df = pd.read_excel(path)

        case_insensetive_columns = [column.lower().strip() for column in df.columns]
        columns = df.columns
        for header in required_headers:
            if not header.lower() in case_insensetive_columns:
                print(f"{header} is required. {header} not is not found in headers")
                raise Exception(f"{header} is required. {header} not is not found in headers")

        for index, row in df.iterrows():
            if row[columns[0]] == "TA" and columns[0] == "GTA":
                new_headers = {i: str(j).strip() for i, j in zip(columns, row)}
                df.rename(columns=new_headers, inplace=True)
                self.data_frame = df
            else:
                rename_headers = {}
                for header in required_headers:
                    header_index = case_insensetive_columns.index(header.lower())
                    rename_headers[columns[header_index]] = header

                df.rename(columns=rename_headers, inplace=True)
                self.data_frame = df

        return self

    def create_student(self):
        student = None
        for index, row in self.data_frame.iterrows():
            if not pd.isna(row["TA Hours"]) and not row["TA Name"] == "TA Name":
                if not pd.isna(row["TA Name"]):
                    student = Student(
                        name=row["TA Name"],
                        email=row["Email"],
                        student_id=row["Student Number"],
                        salary=row["Wage/month"],
                        hours_per_semester="{{Not Specified}}" if pd.isna(row["Total Hours"]) else int(row["Total Hours"]),
                        position=row["Position"]
                    )
                    self.students.append(student)

                assignment = Assignment(
                    subject=row["Subject"],
                    course=int(row["Course Code"]),
                    section=row["Sec No."],
                    type=row["Act Type"],
                    days_met=row["Days Met"],
                    start_time=row["Start Time"],
                    end_time=row["End Time"],
                    hours=int(row["TA Hours"]),
                )
                student.duty_hours += int(row["TA Hours"])
                student.assignments.append(assignment)
                student.assigned_courses[int(row["Course Code"])] = (row["Instructor"], row["Subject"])

        return self

#
# if __name__ == '__main__':
#     ta = TA("data/main_data.xlsx")
#
#     assert 1 == 1
