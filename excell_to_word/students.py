import numpy as np
import pandas as pd


class Student:

    def __init__(self, name, email=None):
        self.name = name
        self.email = email
        self.assignments = []


class Assignment:
    def __init__(self, subject, course, section, type,
                 days_met, start_time, end_time, hours):

        self.subject = subject
        self.course = course
        self.section = section
        self.type = type
        self.start_time = start_time
        self.end_time = end_time
        self.hours = hours


class TA:
    def __init__(self, path):
        self.data_frame = None
        self.read_data(path)
        self.create_student()
        self.students = []

    def read_data(self, path):
        df = pd.read_excel(path)
        colums = df.columns
        for index, row in df.iterrows():
            if row[colums[0]] == "TA" and colums[0] == "GTA":
                new_headers = {i:j for i, j in zip(colums, row)}
                df.rename(columns=new_headers, inplace=True)
                self.data_frame = df

        return self

    def create_student(self):
        for index, row in self.data_frame.iterrows():
            if not pd.isna(row[-1]) and not row["TA"] == "TA":
                if not pd.isna(row["TA"]):
                    student = Student(name=row["TA"], email=row["Email"])
                if not pd.isna(row[-1]):
                    assignment = Assignment(*row[3:])
                    student.assignments.append(assignment)

                self.students.append(student)

        return self


if __name__ == '__main__':
    ta = TA("data/ta_data.xlsx")

    assert 1 == 1
