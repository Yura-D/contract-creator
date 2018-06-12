from docx import Document
from docx.shared import Inches
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime



scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("client_secret.json", scope)
client = gspread.authorize(creds)

sheet = client.open("employee's info").sheet1

title_name = sheet.findall("ПІБ як в паспорті")
if len(title_name) > 1:
    print("There are more than one cell with - ПІБ як в паспорті:\n", title_name)
    print("Please change it in Google sheet")
    quit()
else: 
    title_name = str(title_name)
# checker that title on it place

parts_title_name = title_name.split()
part_pos = parts_title_name[1]
c_position = part_pos.find("C")
number_column_name = part_pos[c_position+1:]


column_names = sheet.col_values(number_column_name)

search = input("Find the emploee: ")

search_results = list()


while True:
    for name in column_names[1:]:  
        if search in name:
            search_results.append(name)
           
    if len(search_results) == 0:
        print("We don't found: ", search)
        del search_results[:]
        search = input("Please try one more time: ")
    elif len(search_results) > 1:
        for result in search_results:
            print("-", result)
        del search_results[:]
        search = input("There is more than one result. Please Try again: ")
    else:
        print("-", search_results[0])
        break

row_number = 0
for name in column_names:
    row_number = row_number + 1
    if name == search_results[0]:
        break


while True:
    contract_date = input("Please type the date of contract - dd.mm.yyyy (Example - 08.05.2018):\n")
    try:
        contract_datetime = datetime.strptime(contract_date, "%d.%m.%Y")
        if contract_datetime.year < 2000 or contract_datetime.year > 2050:
            print("Incorrect format")
            continue
    except ValueError:
        print("Incorrect format")
        continue
    break
# checker to corect date format

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

unit = sheet.row_values(row_number)


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
# You can put above the names of template files and it will work with it


templates_count = 0

templates_dict = dict()

print("\nTemplates: ")
for example in templates:
    templates_count = templates_count + 1
    templates_dict[str(templates_count)] = example
    print(templates_count, "-", example)
# for printing the list of all templates and create the dict with this tampletes and number of choosing

get_file = input("Please choose the files you want to fill (Example: \"1 5 12\"):\n")

choose = get_file.split()

templates_choose = list()

for number_choose in choose:
    templates_choose.append(templates_dict.get(number_choose))
# for choosing templates

if len(unit[8]) < 1:
    fop_address = unit[7]
else:
    fop_address = unit[8]
# for using Address if there not FOP Address

for template in templates_choose:
    file_name_date = contract_date[6:10] + contract_date[2:6] + contract_date[0:2]
    new_file = file_name_date + "-" + str(template)
    # to create the file name with revers position of date

    text_replace("${Contract date}",        date_by_word, template)
    text_replace("${Name}",                 search_results[0], new_file)
    text_replace("${Passport ID}",          unit[3], new_file)
    text_replace("${Passport ID letter}",   unit[3][0:2], new_file)
    text_replace("${Passport ID number}",   unit[3][2:], new_file)
    text_replace("${Passport address}",     unit[4], new_file)
    text_replace("${Passport date}",        unit[6], new_file)
    text_replace("${Address}",              unit[7], new_file)
    text_replace("${FOP address}",          fop_address, new_file)
    text_replace("${ID}",                   unit[10], new_file)
    text_replace("${Bank}",                 unit[11], new_file)
    text_replace("${Bank info}",            unit[13], new_file)
    text_replace("${Person bank info}",     unit[14], new_file)
    text_replace("${FOP ID}",               unit[15], new_file)
    text_replace("${FOP ID date}",          unit[16], new_file)
    text_replace("${Born}",                 unit[20], new_file)
# to replace everything that you need in the template

print("Done")