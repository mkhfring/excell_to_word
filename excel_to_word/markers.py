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

instructor_marking_required_headers = [
    "Instructor",
    "Course",
    "Markers",
    "Email",
    "UBC Email",
    "Marking Hours"

]
instructor_invigilating_required_headers = [
    "Instructor",
    "Course",
    "Markers",
    "Email",
    "UBC Email",
    "Invigilation Hours"
]

class Instructor:
    def __init__(self, name, email, room):
        self.name = name
        self.email = email
        self.room = room
        self.markers = []
        self.duties = []


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
        self.marker_names = {}
        self.instructors = {}

        self.read_data(path)
        self.create_entity()

    def read_data(self, path):

        df = pd.read_excel(path, sheet_name=None)

        if self.role == "student":
            required_headers = ta_required_headers
            self.read_data_sheet(required_headers, df, "ta")


            #     required_headers = ta_required_headers
            #     df = df["TA Invigilating"]
            #     self.read_data_sheet(required_headers, df)
            # if key == "TA Marking":
            #
            #
            # if key == "Instructor Marking":
            #     required_headers = instructor_marking_required_headers
            #     df = df["Instructor Marking"]
            #     self.read_data_sheet(required_headers, df)
            #
            # if key == "Instructor Invigilation":
            #     required_headers = instructor_invigilating_required_headers
            #     df = df["Instructor Invigilation"]
            #     self.read_data_sheet(required_headers, df)

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

        return self

    def create_entity(self):

        for key, value in self.data_frame.items():
            student = None
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
                            self.markers.append(student)


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

        return self

#
# if __name__ == '__main__':
#     ta = TA("data/main_data.xlsx")
#
#     assert 1 == 1
