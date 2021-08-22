import os
from functools import partial

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


def get_template(type):
    print(type)
    if type == "official":
        template_path = filedialog.askopenfilename(
            initialdir=".",
            title="Select official letter template file",
            filetypes=(("template files", "*.docx"), ("all files", "*.*"))
        )
        files_name['official'] = template_path
        lable = tk.Label(frame, text="official letter template file is received", bg="gray")
        lable.pack()

    if type =="offer":
        template_path = filedialog.askopenfilename(
            initialdir=".",
            title="Select offer letter template file",
            filetypes=(("template files", "*.docx"), ("all files", "*.*"))
        )
        files_name['offer'] = template_path
        lable = tk.Label(frame, text="offer leeter template file is received", bg="gray")
        lable.pack()

def get_output():
    out_directory = filedialog.askdirectory()
    files_name["output"] = out_directory


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

official_button = tk.Button(
    frame,
    text="Please Select the official letter template File",
    padx=10,
    pady=5,
    fg='Blue',
    command=partial(get_template, 'official')
)
offer_button = tk.Button(
    frame,
    text="Please Select the offer letter template File",
    padx=10,
    pady=5,
    fg='Blue',
    command=partial(get_template, 'offer')
)
output_button = tk.Button(
    frame,
    text="Please Specify the output directory",
    padx=10,
    pady=5,
    fg='Blue',
    command=get_output
)
data_button.pack()
offer_button.pack()
official_button.pack()
output_button.pack()
root.mainloop()
