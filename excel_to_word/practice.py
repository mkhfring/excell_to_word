from docx import Document


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



input = Document("templates/offer.docx")
out_put = Document()
for paragraph in input.paragraphs:
    if paragraph.text == "":
        continue
    handel_paraghraph(out_put, paragraph)
print(input.sections)

print(dir(input))

out_put.save("output.docx")
