from docx import Document
from docx.shared import Inches

def text_replace(old_text, new_text):
    doc = Document("test.docx")
    for p in doc.paragraphs:
        if old_text in p.text:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if old_text in inline[i].text:
                    text = inline[i].text.replace(old_text, new_text)
                    inline[i].text = text
    doc.save('dest1.docx')

text_replace("Директора", "Дячук Юрій Ігорович")
