from docx import Document
from docx.shared import Inches


template = "test.docx"

contract_date = input("Please type the date of contract (Example - 08.05.2018):\n")

file_name_date = contract_date[6:10] + contract_date[2:6] + contract_date[0:2]
new_file = file_name_date + "-" + str(template)
#to create the file name with revers position of date

month_by_word = str()
if contract_date[3:5] == "01":
    month_by_word = "січня"
elif contract_date[3:5] == "02":
    month_by_word = "лютого"
elif contract_date[3:5] == "03":
    month_by_word = "березня"
elif contract_date[3:5] == "04":
    month_by_word = "квітня"
elif contract_date[3:5] == "05":
    month_by_word = "травня"
elif contract_date[3:5] == "06":
    month_by_word = "червня"
elif contract_date[3:5] == "07":
    month_by_word = "липня"
elif contract_date[3:5] == "08":
    month_by_word = "серпня"
elif contract_date[3:5] == "09":
    month_by_word = "вересня"
elif contract_date[3:5] == "10":
    month_by_word = "жовтня"
elif contract_date[3:5] == "11":
    month_by_word = "листопада"
elif contract_date[3:5] == "12":
    month_by_word = "грудня"

date_by_word = contract_date[0:1] + " " + month_by_word + " " + contract_date[6:10]
#date by word for replacing date instide in contract

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

text_replace("Директора", "Юрій", template)
text_replace("Petro", "Іван", new_file)
text_replace("dfdfdj", "Hello world", new_file)
