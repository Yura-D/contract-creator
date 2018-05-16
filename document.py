from docx import Document
from docx.shared import Inches


template = "test.docx"

contract_date = input("Please type the date of contract (Example - 08.05.2018):\n")

file_name_date = contract_date[6:10] + contract_date[2:6] + contract_date[0:2]
new_file = file_name_date + "-" + str(template)

def text_replace(old_text, new_text, file):
    doc = Document(file)
    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for line in inline:
                if old_text in line.text:
                    text = line.text.replace(old_text, new_text)
                    line.text = text
    doc.save(new_file)

text_replace("Директора", "Дячук Юрій Ігорович", template)
text_replace("Petro", "Іван", new_file)
text_replace("dfdfdj", "Hello world", new_file)
