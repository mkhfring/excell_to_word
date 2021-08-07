import os
import glob

import docx
import click

from excel_to_word.tempelate import SentenceElement, template
from excel_to_word.students import TA


HERE = os.path.dirname(os.path.realpath(__file__))
paragraphs = []


def create_first_paragraph(document, name):
    paragraph1 = f"""
Dear {name},

Thank you very much for applying for our Teaching Assistant (TA) positions. \
We are in the process of finalizing the TA assignments for the first term of the 2021 Winter semester. \
I have the following TA offer for you:  """
    document.add_paragraph(paragraph1)


def create_template(document, element):
    if element[1] == SentenceElement.PARAGRAPH:
        p = document.add_paragraph(element[0])
        paragraphs.append(p)

    if element[1] == SentenceElement.BOLD_PARAGRAPH:
        p = document.add_paragraph()
        p.add_run(element[0]).bold = True
        paragraphs.append(p)

    if element[1] == SentenceElement.SENTENCE:
        p = paragraphs[-1]
        p.add_run(element[0])

    if element[1] == SentenceElement.BOLD:
        p = paragraphs[-1]
        p.add_run(element[0]).bold = True

    if element[1] == SentenceElement.UNDER_LINE:
        p = paragraphs[-1]
        p.add_run(element[0]).underline = True


def add_table(document, student, headers):
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


def create_second_paragraph(document, content):
    for element in content:
        create_template(document, element)


def handel_paraghraph(document, paragraph):
    para = document.add_paragraph()
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


@click.command()
@click.option('--path', default=os.path.join(HERE, "data/test.xlsx"),
              help='number of greetings')
def main(path):
    print(path)
    ta = TA(os.path.join(path))
    files = glob.glob(os.path.join(HERE, 'data/*'))
    for f in files:
        if f.endswith(".xlsx"):
            continue

        os.remove(f)
    for student in ta.students:
        doc = docx.Document()
        input = docx.Document(os.path.join(HERE, "templates/offer.docx"))
        for paragraph in input.paragraphs:
            if paragraph.text == "":
                continue
            if paragraph.text.lower() == "table":
                add_table(doc, student, ta.data_frame.columns[3:])
                continue

            handel_paraghraph(doc, paragraph)

        # create_first_paragraph(doc, student.name)
        # add_table(doc, student, ta.data_frame.columns[3:])
        # create_second_paragraph(doc, template)
        doc.save(os.path.join(HERE, f"data/{student.name}.docx"))


if __name__ == "__main__":
    main()