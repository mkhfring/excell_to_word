import os
import docx


print("Initiating the code")

doc = docx.Document()
doc.add_heading("This is a test document", 0)

fack_data = [
    [1, "first record", "second recod"],
    [2, "third record", "fourth recod"]
]
some_table = doc.add_table(1, 3)
some_table.style = "Table Grid"
for id, element1, element2 in fack_data:
    row_cells = some_table.add_row().cells
    row_cells[0].text = str(id)
    row_cells[1].text = element1
    row_cells[2].text = element2

doc.save("test_doc.docx")
os.system("start test_doc.docx")
