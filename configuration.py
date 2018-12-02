secret = "client_secret.json"
# for google sheets
# need to create in google developer console

sheet_name = "employee's info"
sheet_list = 0
# for sheet where you get info

client_id = 'client_id.json'
# client for google drive - token_id. 
# You need to create in google developer console like 
# OAuth 2.0 client IDs for Google Drive

drive_token = 'token.json'
# token for google drive. Created automatically

templates_folder = "1o74052-am8082dMqhlY3TYBztUb_BO2s"
# place where you hold templates

personal_data_folder = "1zJNVEIXn897T7bNgdfIyYN08SC550-Ur"
contract_folder = "1wXp-05T4PAMwh_t0eN8C30zabruXT2GC"
NDA_folder = "1vKxzyS47ZCBaDbsTWQXIYu3nFDg1jDJg"
# folders for uploading contracts


folder_dict = {
    # for uploading in different folders
    "Test_1.docx": personal_data_folder,
    "Test_2.docx": contract_folder,
    "Test_3.docx": NDA_folder
}

# register

register_sheet_red = "1l_5mXcxuCwV_NWsl3o21Bd4Fq7B_PN5ikg5vQfmVkN8"
red_sheet_list = 0
register_sheet_blue = "1x9yg5Ma1hD2ZnVTrwUn9u2g9cZaNuJToWdAgaCDr7dQ"
blue_sheet_list = 0 
register_sheet_green = "18R9ZnXsHJiQlbTRHUA5zE10yXt_Q6tcjgYaAF_voPXs"
green_sheet_list = 0 


register_dict = {
    "Test_1.docx": "register_green",
    "Test_2.docx": "register_blue",
    "Test_3.docx": "register_red"
}
