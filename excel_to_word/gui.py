import os

import tkinter as tk
from tkinter import filedialog, Text

files_name = {}


def get_data():
    data_path = filedialog.askopenfilename(
        initialdir=".",
        title="Select excel file",
        filetypes=(("excel files", "*.xlsx"), ("all files", "*.*"))
    )
    files_name['input'] = data_path
    lable = tk.Label(frame, text="Excel file is received", bg="gray")
    lable.pack()


def get_template():
    template_path = filedialog.askopenfilename(
        initialdir=".",
        title="Select excel file",
        filetypes=(("template files", "*.docx"), ("all files", "*.*"))
    )
    files_name['template'] = template_path
    lable = tk.Label(frame, text="template file is received", bg="gray")
    lable.pack()


root = tk.Tk()
#canvas = tk.Canvas(root, height=400, width=400, bg="green")
#canvas.pack()
frame = tk.Frame(root, bg="white")
frame.place(width=500, height=500, relx=0.4, rely=0.1)
data_button = tk.Button(
    frame,
    text="Please Select the Excel File",
    padx=10,
    pady=5,
    fg='Blue',
    command=get_data
)

template_button = tk.Button(
    frame,
    text="Please Select the template File",
    padx=10,
    pady=5,
    fg='Blue',
    command=get_template
)
data_button.pack()
template_button.pack()
root.mainloop()