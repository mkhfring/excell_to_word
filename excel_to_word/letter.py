import copy
import glob
import os
import re

import docx
from docx.shared import Pt

from excel_to_word.students import TA


HERE = os.path.dirname(os.path.realpath(__file__))


class Letter:
    def __init__(self, data_path, template_path, output_template_path=None):
        self.output_template_path = output_template_path
        template = docx.Document(template_path)
        template_paragraphs = [paragraph for paragraph in template.paragraphs]
        self.paragraphs = template_paragraphs
        self.ta_assignment = TA(data_path)

    def add_table(self, document, student, headers):
        document.add_paragraph(f" {student.duty_hours} hrs/wk for following:")
        assignments_headers = student.assignments[0].__dict__.keys()
        some_table = document.add_table(1, len(headers))
        some_table.style = "Table Grid"
        first_row_cells = some_table.rows[0].cells
        for index, header in enumerate(headers):
            if header in ["TA", "Student No.", "Email"]:
                continue
            first_row_cells[index].text = str(header)

        for index, assignment in enumerate(student.assignments):
            nex_cell = some_table.add_row().cells
            for index, header in enumerate(assignments_headers):
                assert 1 == 1
                nex_cell[index].text = str(getattr(assignment, header))

        document.add_paragraph()
        return self

    def handel_paragraphs(self, document, paragraph, type=0):
        para = document.add_paragraph("")
        for run in paragraph.runs:
            output_run = para.add_run(run.text)
            output_run.bold = run.bold
            output_run.italic = run.italic
            output_run.underline = run.underline
            output_run.font.color.rgb = run.font.color.rgb
            # Run's font data
            output_run.style.name = run.style.name
            # Paragraph's alignment data
        para.paragraph_format.alignment = paragraph.paragraph_format.alignment
        if type:
            style = document.styles['Normal']
            #para.alignment = 0
            font = style.font
            font.name = "Arial"
            font.size = Pt(9)
            para.style = document.styles['Normal']

        return self

    def replace_token(self, paragraph, student):
        NotImplemented

    def create_file(self, student):
        NotImplemented

    def create_output(self):
        for student in self.ta_assignment.students:
            self.create_file(student)



class OfferLetter(Letter):
    def __init__(self, data_path,  template_path):
        super().__init__(data_path, template_path)

    def create_file(self, student):
        doc = docx.Document()
        for paragraph in self.paragraphs:
            if paragraph.text == "":
                continue
            if paragraph.text == "\n":
                continue

            if paragraph.text.lower() == "table":
                self.add_table(doc, student, self.ta_assignment.data_frame.columns[3:])
                continue

            paragraph = self.replace_token(paragraph, student)
            self.handel_paragraphs(doc, paragraph, type=1)

        doc.save(os.path.join(HERE, f"data/{student.name}.docx"))
        print(f"The document for {student.name} is saved")

    def replace_token(self, paraghraph, student):
        new_paragraph = copy.deepcopy(paraghraph)
        name_pattern = r"\[first_name\]"
        if re.search(name_pattern, new_paragraph.text):
            replaced_text = re.sub(name_pattern, student.name.split(" ")[0], new_paragraph.text)
            new_paragraph.text = replaced_text

        return new_paragraph


class OfficialLetter(Letter):
    def __init__(self, data_path, template_path, output_template_path):
        super().__init__(data_path, template_path, output_template_path)

    def create_file(self, student):
        official_letters_path = os.path.join(HERE, "data/OfficialLetters")
        files = glob.glob(official_letters_path)
        for f in files:
            if os.path.isdir(f):
                continue

            if f.endswith(".xlsx"):
                continue

            os.remove(f)

        doc = docx.Document(self.output_template_path)
        self.replace_footer(doc, student)
        for paragraph in self.paragraphs:
                # if paragraph.text == "":
                #     continue
                # if paragraph.text == "\n":
                #     continue

            if re.search("\[Table_\]", paragraph.text):
                self.add_table(doc, student, self.ta_assignment.data_frame.columns[3:])
                continue
            paragraph = self.replace_token(paragraph, student)
            self.handel_paragraphs(doc, paragraph, type=1)

        doc.save(os.path.join(HERE, f"data/OfficialLetters/{student.name}.docx"))
        print(f"The document for {student.name} is saved")

    def replace_token(self, paraghraph, student):

        new_paragraph = copy.deepcopy(paraghraph)
        firstname_pattern = r"\[First_Name_\]"
        lastname_pattern = r"\[Last_Name_\]"
        email_pattern = r"\[Email_\]"
        hours_per_week_pattern = r"\[Hours_Per_Week_\]"
        student_name = student.name.split(" ")

        if re.search(firstname_pattern, new_paragraph.text):
            replaced_text = re.sub(firstname_pattern, student.name.split(" ")[0], new_paragraph.text)
            new_paragraph.text = replaced_text

        if re.search(lastname_pattern, new_paragraph.text):
            replaced_text = re.sub(
                lastname_pattern,
                student_name[1] if len(student_name) > 1 else student_name[0],
                new_paragraph.text
            )
            new_paragraph.text = replaced_text

        if re.search(email_pattern, new_paragraph.text):
            replaced_text = re.sub(
                email_pattern,
                student.email,
                new_paragraph.text
            )
            new_paragraph.text = replaced_text

        if re.search(hours_per_week_pattern, new_paragraph.text):
            replaced_text = re.sub(
                hours_per_week_pattern,
                str(student.duty_hours),
                new_paragraph.text
            )
            new_paragraph.text = replaced_text

        return new_paragraph

    def replace_footer(self, doc, student):
        footer = doc.sections[0].footer
        name = student.name.split(" ")
        for paragraph in footer.paragraphs:
            if re.search("\[First_Name_\]", paragraph.text):
                replaced_text = re.sub(
                    "\[Student_ID\]",
                    str(student.student_id),
                    re.sub(
                        "\[Last_Name_\]",
                        name[1] if len(name)>1 else name[0],
                        re.sub(
                            "\[First_Name_\]",
                            student.name.split(" ")[0],
                            paragraph.text
                        )
                    )
                )
                paragraph.text = replaced_text


if __name__ == '__main__':

    files = glob.glob(os.path.join(HERE, 'data/*'))
    for f in files:
        if os.path.isdir(f):
            continue

        if f.endswith(".xlsx"):
            continue

        os.remove(f)

    # input = docx.Document(os.path.join(HERE, "templates/offer.docx"))
    # input_paragraphs = [paragraph for paragraph in input.paragraphs]
    letter = OfficialLetter(
        os.path.join(HERE, "data/test.xlsx"),
        os.path.join(HERE, "templates/letters.docx"),
        os.path.join(HERE, "templates/letters_temp.docx")
    )
    letter.create_output()
    # letter = OfferLetter(
    #     os.path.join(HERE, "data/test.xlsx"),
    #     os.path.join(HERE, "templates/offer.docx"),
    # )
    # letter.create_output()
    assert 1 == 1
