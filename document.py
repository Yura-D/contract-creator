from docx import Document
from docx.shared import Inches


#document = Document("test.docx")
Director = "Дячук Юрій Ігорович"

doc = Document("test.docx")
for p in doc.paragraphs:
    if 'Директора' in p.text:
        inline = p.runs
        # Loop added to work with runs (strings with same style)
        for i in range(len(inline)):
            if 'Директора' in inline[i].text:
                text = inline[i].text.replace('Директора', Director)
                inline[i].text = text
        print (p.text)

doc.save('dest1.docx')
