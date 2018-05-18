from docx import Document
from docx.shared import Inches


template = "test.docx"

contract_date = input("Please type the date of contract (Example - 08.05.2018):\n")

file_name_date = contract_date[6:10] + contract_date[2:6] + contract_date[0:2]
new_file = file_name_date + "-" + str(template)
#to create the file name with revers position of date

def get_month(contract_date):
    return {
        "01": "січня",
        "02": "лютого",
        "03": "березня",
        "04": "квітня",
        "05": "травня",
        "06": "червня",
        "07": "липня",
        "08": "серпня",
        "09": "вересня",
        "10": "жовтня",
        "11": "листопада",
        "12": "грудня"
    }.get(contract_date[3:5])

month_by_word = get_month(contract_date)

date_by_word = contract_date[0:2] + " " + month_by_word + " " + contract_date[6:10]
#date by word for replacing date instide in contract

c_name = "test1"
c_passport_id = "test2"
c_passport_address = "test3"
c_passport_date = "test4"
c_address = "test5"
c_fop_address = "test6"
c_id = "test7"
c_bank = "test8"
c_bank_info = "test9"
c_person_bank_info = "test10"
c_fop_id = "test11"
c_fop_id_date = "test12"
c_born = "test13"
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

