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

        self.datachech = tk.Checkbutton(self.root, text="Data is attached")
        self.datachech.place(x=620, y=1)

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
                self.datachech.select()

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
