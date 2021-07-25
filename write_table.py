import os
import docx


print("Initiating the code")

doc = docx.Document()
doc.add_heading("This is a test document", 0)
doc.save("test_doc.docx")
os.system("start test_doc.docx")
