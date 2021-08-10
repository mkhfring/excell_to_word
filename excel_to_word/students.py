import numpy as np
import pandas as pd


class Student:

    def __init__(self, name, student_id, email=None):
        self.name = name
        self.email = email
        self.student_id = student_id
        self.assignments = []
        self.duty_hours = 0


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
        colums = df.columns
        for index, row in df.iterrows():
            if row[colums[0]] == "TA" and colums[0] == "GTA":
                new_headers = {i: str(j).strip() for i, j in zip(colums, row)}
                df.rename(columns=new_headers, inplace=True)
                self.data_frame = df

        return self

    def create_student(self):
        student = None
        for index, row in self.data_frame.iterrows():
            if not pd.isna(row[-1]) and not row["TA"] == "TA":
                if not pd.isna(row["TA"]):
                    student = Student(name=row["TA"], email=row["Email"], student_id=row["Student no."])
                    self.students.append(student)

                assignment = Assignment(
                    subject=row["Subject"],
                    course=row["Course"],
                    section=row["Sec No."],
                    type=row["Act. Type"],
                    days_met=row["Days Met"],
                    start_time=row["Start time"],
                    end_time=row["End time"],
                    hours=row["TA Hours"],
                )
                student.duty_hours += int(row["TA Hours"])
                student.assignments.append(assignment)

        return self


if __name__ == '__main__':
    ta = TA("data/test.xlsx")

    assert 1 == 1
