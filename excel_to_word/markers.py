import numpy as np
import pandas as pd


ta_required_headers = [
    "Name",
    "Email",
    "Course",
    "Room",
    "Instructor",
    "Instructor Email",
    "Date",
    "Time",
    "No. Hours"
]

instructor_required_headers = { "marking" : [
    "Instructor",
    "Course",
    "Markers",
    "Email",
    "UBC Email",
    "Marking Hours"

],
    "invigilation":
[
    "Instructor",
    "Course",
    "Markers",
    "Email",
    "UBC Email",
    "Invigilation Hours"
]}


class Instructor:
    def __init__(self, name):
        self.name = name
        self.marking_courses = []
        self.invigilation_courses = []
        self.marking_course_names = {}
        self.invigilating_course_names = {}


class Course:
    def __init__(self, course, ):
        self.course = course
        self.assignments = []


class CourseAssignment:
    def __init__(self, markers, email, ubc_email, invigilation_hour):
        self.markers = markers
        self.email = email
        self.ubc_email = ubc_email
        self.invigilation_hour = invigilation_hour


class Duty:
    def __init__(self, course, date, time, n_hours, instructor, instructor_email, room):
        self.course = course
        self.date = date
        self.time = time
        self.instructor = instructor
        self.instructor_email = instructor_email
        self.room = room
        self.n_hours = n_hours


class Marker:

    def __init__(self, name, email):

        self.name = name
        self.email = email
        self.assignments = []
        self.courses = []
        self.instructor_names = []
        self.instructors = []
        self.marking_duties = []
        self.invigilation_duties = []


class TA:
    def __init__(self, path, role, duty=None):
        self.data_frame = {}
        self.role = role
        self.duty = duty
        self.markers = list()
        self.teachers = list()
        self.marker_names = {}
        self.teacher_names = {}
        self.instructors = {}

        self.read_data(path)
        self.create_entity()

    def read_data(self, path):

        df = pd.read_excel(path, sheet_name=None)

        if self.role == "student":
            required_headers = ta_required_headers
            self.read_data_sheet(required_headers, df, "ta")

        else:
            required_headers = instructor_required_headers
            self.read_data_sheet(required_headers, df, "teacher")

    def read_data_sheet(self, required_headers, df, role_duty):
        if role_duty == "ta":
            for sheet in ["TA Invigilating", "TA Marking"]:
                df1 = df[sheet]
                case_insensetive_columns = [column.lower() for column in df1.columns]
                columns = df1.columns
                for header in required_headers:
                    if not header.lower() in case_insensetive_columns:
                        print(f"{header} is required. {header} not is not found in headers")
                        raise Exception(f"{header} is required. {header} not is not found in headers")

                for index, row in df1.iterrows():
                    if row[columns[0]] == "TA" and columns[0] == "GTA":
                        new_headers = {i: str(j).strip() for i, j in zip(columns, row)}
                        df1.rename(columns=new_headers, inplace=True)
                        self.data_frame[sheet] = df1
                    else:
                        rename_headers = {}
                        for header in required_headers:
                            header_index = case_insensetive_columns.index(header.lower())
                            rename_headers[columns[header_index]] = header

                        df1.rename(columns=rename_headers, inplace=True)
                        self.data_frame[sheet] = df1
        else:
            for sheet in ["Instructor Invigilation", "Instructor Marking"]:
                df1 = df[sheet]
                case_insensetive_columns = [column.lower().strip() for column in df1.columns]
                columns = df1.columns
                if "marking" in sheet.lower():
                    extracted_required_headers = required_headers["marking"]
                else:
                    extracted_required_headers = required_headers["invigilation"]

                for header in extracted_required_headers:
                    if not header.lower() in case_insensetive_columns:
                        print(f"{header} is required. {header} not is not found in headers")
                        raise Exception(f"{header} is required. {header} not is not found in headers")

                for index, row in df1.iterrows():
                    if row[columns[0]] == "TA" and columns[0] == "GTA":
                        new_headers = {i: str(j).strip() for i, j in zip(columns, row)}
                        df1.rename(columns=new_headers, inplace=True)
                        self.data_frame[sheet] = df1
                    else:
                        rename_headers = {}
                        for header in extracted_required_headers:
                            header_index = case_insensetive_columns.index(header.lower())
                            rename_headers[columns[header_index]] = header

                        df1.rename(columns=rename_headers, inplace=True)
                        self.data_frame[sheet] = df1

        return self

    def create_entity(self):

        for key, value in self.data_frame.items():
            student = None
            teacher = None
            course = None
            for index, row in value.iterrows():
                if self.role == "student":

                    if not pd.isna(row["Name"]):
                        if not row["Name"] in self.marker_names:
                            student = Marker(
                                name=row["Name"],
                                email=row["Email"],
                            )
                            self.marker_names[row["Name"]] = student
                            self.markers.append(student)
                        else:
                            student = self.marker_names[row["Name"]]
                            #self.markers.append(student)


                    if pd.isna(row["Date"]):
                        continue

                    assignment = Duty(
                        course=row["Course"],
                        date=row["Date"],
                        time=row["Time"],
                        n_hours=row["No. Hours"],
                        instructor=row["Instructor"],
                        instructor_email=row["Instructor Email"],
                        room=row["Room"]
                    )
                    if "marking" in key.lower() and student is not None:
                        student.marking_duties.append(assignment)
                    if "invigilating" in key.lower() and student is not None:
                        try:
                            student.invigilation_duties.append(assignment)
                        except:
                            pass
                else:

                    if not pd.isna(row["Instructor"]):
                        if not row["Instructor"] in self.teacher_names:
                            teacher = Instructor(
                                name=row["Instructor"]
                            )
                            self.teacher_names[row["Instructor"]] = teacher
                            self.teachers.append(teacher)
                        else:
                            teacher = self.teacher_names[row["Instructor"]]

                    if pd.isna(row["Markers"]):
                        continue

                    if not pd.isna(row["Course"]):
                        course = Course(
                            course=row["Course"],

                        )

                    if "marking" in key.lower() and course is not None:
                        cours_assingment = CourseAssignment(
                            markers=row["Markers"],
                            email=row["Email"],
                            ubc_email=row["UBC Email"],
                            invigilation_hour=row["Marking Hours"]
                        )
                        course.assignments.append(cours_assingment)
                        if course.course not in teacher.marking_course_names:
                            teacher.marking_courses.append(course)
                            teacher.marking_course_names[row['Course']] = course

                    if "invigilation" in key.lower() and course is not None:
                        cours_assingment = CourseAssignment(
                            markers=row["Markers"],
                            email=row["Email"],
                            ubc_email=row["UBC Email"],
                            invigilation_hour=row["Invigilation Hours"]
                        )
                        course.assignments.append(cours_assingment)
                        try:
                            if course.course not in teacher.invigilating_course_names:
                                teacher.invigilation_courses.append(course)
                                teacher.invigilating_course_names[row['Course']] = course

                        except:
                            pass

        return self

#
# if __name__ == '__main__':
#     ta = TA("data/main_data.xlsx")
#
#     assert 1 == 1
