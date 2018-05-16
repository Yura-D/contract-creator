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

date_by_word = contract_date[0:2] + " " + month_by_word + " " + contract_date[6:10]
#date by word for replacing date instide in contract

c_name = "1"
c_passport_id = "2"
c_passport_address = "3"
c_passport_date = "4"
c_address = "5"
c_fop_address = "6"
c_id = "7"
c_bank = "8"
c_bank_info = "9"
c_person_bank_info = "10"
c_fop_id = "11"
c_fop_id_date = "12"
c_born = "13"
#need for future functional. Now only like test

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

text_replace("${Contract date}", date_by_word, template)
text_replace("${Name}", c_name, new_file)
text_replace("${Passport ID}", c_passport_id, new_file)
text_replace("${Passport address}", c_passport_address, new_file)
text_replace("${Passport date}", c_passport_date, new_file)
text_replace("${Address}", c_address, new_file)
text_replace("${FOP address}", c_fop_address, new_file)
text_replace("${ID}", c_id, new_file)
text_replace("${Bank}", c_bank, new_file)
text_replace("${Bank info}", c_bank_info, new_file)
text_replace("${Person bank info}", c_person_bank_info, new_file)
text_replace("${FOP ID}", c_fop_id, new_file)
text_replace("${FOP ID date}", c_fop_id_date, new_file)
text_replace("${Born}", c_born, new_file)
#to replace everything that you need in the template
