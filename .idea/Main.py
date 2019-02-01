import gspread
import pprint as pp#this one is needed to format correctly the data, otherwise it is not legible
from oauth2client.service_account import ServiceAccountCredentials

# credentials to create a client to interact with the GDrive API
scope = ["https://spreadsheets.google.com/feeds",
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name("FinancialModeling.json", scope)
client = gspread.authorize(creds)

# adding the workbook by name
# double check the name of the excel file is correct or it will return error

sheet = client.open("BaseModel").sheet1  # returns weird values from sheet CP (control panel) with checks - TRY TO UNDERSTAND HOW TO USE THIS
values = sheet.get_all_records()
pp.pprint(values)

