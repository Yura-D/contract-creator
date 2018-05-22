from docx import Document
from docx.shared import Inches
import gspread
from oauth2client.service_account import ServiceAccountCredentials



scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("client_secret.json", scope)
client = gspread.authorize(creds)

sheet = client.open("employee's info").sheet1

search = input("Find the emploee (enter full name): ")


unit = sheet.findall(search)
unit = str(unit)
row_unit = unit.split("[<Cell R")
row_unit = row_unit[1].split("C")
row_number = row_unit[0]
# make better pursing of row number

contract_date = input("Please type the date of contract (Example - 08.05.2018):\n")

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
# date by word for replacing date instide in contract

c_name = sheet.cell(row_number, 2).value
c_passport_id = sheet.cell(row_number, 4).value
c_passport_address = sheet.cell(row_number, 5).value
c_passport_date = sheet.cell(row_number, 7).value
c_address = sheet.cell(row_number, 8).value
c_fop_address = sheet.cell(row_number, 9).value
c_id = sheet.cell(row_number, 11).value
c_bank = sheet.cell(row_number, 12).value
c_bank_info = sheet.cell(row_number, 14).value
c_person_bank_info = sheet.cell(row_number, 15).value
c_fop_id = sheet.cell(row_number, 16).value
c_fop_id_date = sheet.cell(row_number, 17).value
c_born = sheet.cell(row_number, 21).value


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


templates = ["test.docx", "Згода на обробку персональних даних.docx"]

for template in templates:
    file_name_date = contract_date[6:10] + contract_date[2:6] + contract_date[0:2]
    new_file = file_name_date + "-" + str(template)
    # to create the file name with revers position of date





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
    # to replace everything that you need in the template