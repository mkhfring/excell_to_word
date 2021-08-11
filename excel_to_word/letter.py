import copy
import glob
import os
import re

import docx

from excel_to_word.students import TA


HERE = os.path.dirname(os.path.realpath(__file__))


class Letter:
    def __init__(self, data_path, template_path):
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

    def handel_paragraphs(self, document, paragraph):
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

    def replace_token(self, paragraph, student):
        NotImplemented

    def create_output(self):
        NotImplemented


class OfferLetter(Letter):
    def __init__(self, data_path,  template_path):
        super().__init__(data_path, template_path)

    def create_output(self):
        for student in self.ta_assignment.students:
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
                self.handel_paragraphs(doc, paragraph)

            doc.save(os.path.join(HERE, f"data/{student.name}.docx"))

    def replace_token(self, paraghraph, student):

        new_paragraph = copy.deepcopy(paraghraph)
        name_pattern = r"\[first_name\]"
        if re.search(name_pattern, new_paragraph.text):
            replaced_text = re.sub(name_pattern, student.name.split(" ")[0], new_paragraph.text)
            new_paragraph.text = replaced_text

        return new_paragraph


class OfficialLetter(Letter):
    def __init__(self, data_path, paragraphs):
        super().__init__(data_path, paragraphs)

    def create_output(self):
        for student in self.ta_assignment.students:
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
                self.handel_paragraphs(doc, paragraph)

            doc.save(os.path.join(HERE, f"data/{student.name}.docx"))

    def replace_token(self, paraghraph, student):

        new_paragraph = copy.deepcopy(paraghraph)
        name_pattern = r"\[first_name\]"
        if re.search(name_pattern, new_paragraph.text):
            replaced_text = re.sub(name_pattern, student.name.split(" ")[0], new_paragraph.text)
            new_paragraph.text = replaced_text

        return new_paragraph

    def handel_paraghraph(document, paragraph):
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


if __name__ == '__main__':

    files = glob.glob(os.path.join(HERE, 'data/*'))
    for f in files:
        if f.endswith(".xlsx"):
            continue

        os.remove(f)

    # input = docx.Document(os.path.join(HERE, "templates/offer.docx"))
    # input_paragraphs = [paragraph for paragraph in input.paragraphs]
    letter = OfficialLetter(
        os.path.join(HERE, "data/test.xlsx"),
        os.path.join(HERE, "templates/letters.docx")
    )
    letter.create_output()
    assert 1 == 1
