import gspread
import pprint as pp#this one is needed to format correctly the data, otherwise it is not legible
from oauth2client.service_account import ServiceAccountCredentials

# add credentials - create client to interact with GDrive and GSheets API
# renamed credentials and gc after the documentation for correct implementation of the library

scope = ["https://spreadsheets.google.com/feeds",
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name("FinancialModeling.json", scope)
gc = gspread.authorize(credentials)

# adding the workbook by name
# double check the name of the excel file is correct or it will return error

sh = gc.open("BaseModel")  # loading the excel file

F1Y = sh.worksheet("F1Y")  # this one opens the actual worksheet in the spreadsheet
# val = F1Y.acell("A20").value  # select a specific cell in the worksheet
# print(val)  # printing the value just to check (should be "Ammortamenti") = OK!

# now we have to try and edit a cell in the actual excel file, and se if it updates the value

data = sh.worksheet("DATA")
# pp.pprint(data.range("B3:D17"))  test printing a range in the data worksheet = OK!

data.update_acell("D5", 100000)  # editing the value of a cell in sheet data
pp.pprint(data.acell("D5"))  # print to check the value, should return 100k
