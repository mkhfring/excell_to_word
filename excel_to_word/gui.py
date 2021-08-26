import os
import threading
from functools import partial

import tkinter as tk
from tkinter import filedialog, Text

from excel_to_word.letter import OfficialLetter, OfferLetter

files_name = {}
HERE = os.path.dirname(os.path.realpath(__file__))


class ExcelToWordGui:
    def __init__(self):
        self.input_files = {}
        self.root = tk.Tk()
        self.text = tk.StringVar()
        self.text.set("Test")
        self.label = tk.Label(self.root, textvariable=self.text)

        self.button = tk.Button(self.root,
                                text="Click to change text below",
                                command=self.change_text)


        # c = tk.Checkbutton(self.root, text="accept T&C")
        # c.pack()

        self.data_button = tk.Button(
            self.root,
            text="Please Select the Excel File",
            command=partial(self.get_file, "data")
        )
        self.data_button.pack()
        self.offer_button = tk.Button(
            self.root,
            text="Please Select the offer letter template File",
            command=partial(self.get_file, "offer")
        )
        self.offer_button.pack()
        self.official_button = tk.Button(
            self.root,
            text="Please Select the official letter template File",
            command=partial(self.get_file, "official")
        )
        self.official_button.pack()
        self.output_button = tk.Button(
            self.root,
            text="Please Specify the output directory",
            command=partial(self.get_file, "output")
        )
        self.run_button = tk.Button(
            self.root,
            text="Run the project",
            command=self.run_project

        )
        self.run_button.pack()
        self.button.pack()
        self.label.pack()
        self.root.mainloop()

    def run_project(self):
        self.text.set("Start the process of creating output files")
        try:
            data = self.input_files["data"]
        except Exception as e:
            self.text.set("The excel file is not attached")
            raise e

        if self.input_files.get("offer"):
            offer_letter = OfferLetter(
                self.input_files["data"],
                self.input_files["offer"],
            )
            t = threading.Thread(target=offer_letter.create_output)
            t.start()
            self.text.set("output files are created")
        if self.input_files.get("official"):
            official_letter = OfficialLetter(
                self.input_files["data"],
                self.input_files["official"],
                os.path.join(HERE, "templates/letters_temp.docx")
            )
            t1 = threading.Thread(target=official_letter.create_output)
            t1.start()


    def change_text(self):
        self.text.set("Text updated")

    def get_file(self, type):
        if type == "data":
            self.input_files['data'] = filedialog.askopenfilename(
                initialdir=".",
                title="Select excel file",
                filetypes=(("excel files", "*.xlsx"), ("all files", "*.*"))
            )
            if os.path.isfile(self.input_files['data']):
                self.text.set("The excel data is inserted")

        if type == "offer":
            self.input_files['offer'] = filedialog.askopenfilename(
                initialdir=".",
                title="Select offer template file",
                filetypes=(("excel files", "*.docx"), ("all files", "*.*"))
            )
            if os.path.isfile(self.input_files['offer']):
                self.text.set("The offer letter template file is inserted")

        if type == "official":
            self.input_files['official'] = filedialog.askopenfilename(
                initialdir=".",
                title="Select offer template file",
                filetypes=(("excel files", "*.docx"), ("all files", "*.*"))
            )
            if os.path.isfile(self.input_files['official']):
                self.text.set("The official letter template file is inserted")




app = ExcelToWordGui()
# def get_data():
#     data_path = filedialog.askopenfilename(
#         initialdir=".",
#         title="Select excel file",
#         filetypes=(("excel files", "*.xlsx"), ("all files", "*/*"))
#     )
#     files_name['input'] = data_path
#     lable = tk.Label(frame, text="Excel file is received", bg="gray")
#     lable.pack()
#
#
# def get_template(type):
#     print(type)
#     if type == "official":
#         template_path = filedialog.askopenfilename(
#             initialdir=".",
#             title="Select official letter template file",
#             filetypes=(("template files", "*.docx"), ("all files", "*/*"))
#         )
#         files_name['official'] = template_path
#         lable = tk.Label(frame, text="official letter template file is received", bg="gray")
#         lable.pack()
#
#     if type =="offer":
#         template_path = filedialog.askopenfilename(
#             initialdir=".",
#             title="Select offer letter template file",
#             filetypes=(("template files", "*.docx"), ("all files", "*/*"))
#         )
#         files_name['offer'] = template_path
#         lable = tk.Label(frame, text="offer leeter template file is received", bg="gray")
#         lable.pack()
#
#
# def get_output():
#     out_directory = filedialog.askdirectory()
#     files_name["output"] = out_directory
#
#
# def run_project():
#     # official_letter = OfficialLetter(
#     #     files_name['input'],
#     #     os.path.join(HERE, "templates/letters_temp.docx"),
#     #     os.path.join(HERE, "templates/letters_temp.docx")
#     # )
#     # t1 = threading.Thread(target=official_letter.create_output)
#     # t1.start()
#     #official_letter.create_output()
#     offer_letter = OfferLetter(
#         files_name['input'],
#         files_name['offer'],
#     )
#     # t = threading.Thread(target=offer_letter.create_output)
#     # t.start()
#     lable = tk.Label(frame, text="Please wait while the code generating output", bg="gray")
#     lable.pack()
#     offer_letter.create_output()
#     lable2 = tk.Label(frame, text="Output has been generated", bg="gray")
#     lable2.pack()


# root = tk.Tk()
# #canvas = tk.Canvas(root, height=400, width=400, bg="green")
# #canvas.pack()
# frame = tk.Frame(root, bg="white")
# frame.place(width=500, height=500, relx=0.4, rely=0.1)
# data_button = tk.Button(
#     frame,
#     text="Please Select the Excel File",
#     padx=10,
#     pady=5,
#     fg='Blue',
#     command=get_data
# )
#
# official_button = tk.Button(
#     frame,
#     text="Please Select the official letter template File",
#     padx=10,
#     pady=5,
#     fg='Blue',
#     command=partial(get_template, 'official')
# )
# offer_button = tk.Button(
#     frame,
#     text="Please Select the offer letter template File",
#     padx=10,
#     pady=5,
#     fg='Blue',
#     command=partial(get_template, 'offer')
# )
# output_button = tk.Button(
#     frame,
#     text="Please Specify the output directory",
#     padx=10,
#     pady=5,
#     fg='Blue',
#     command=get_output
# )
# run_button = tk.Button(
#     frame,
#     text="Run the project",
#     padx=10,
#     pady=5,
#     fg='Blue',
#     command=run_project
# )
# data_button.pack()
# offer_button.pack()
# official_button.pack()
# output_button.pack()
# run_button.pack()
# root.mainloop()
